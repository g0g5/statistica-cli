using System.CommandLine;
using System.Reflection;

namespace StatisticaCli;

internal static class SkillInstaller
{
    internal static void Add(RootCommand root)
    {
        var command = new Command("install-skill", "Install the bundled skill and guide.");
        var global = new Option<bool>("--global") { Description = "Install under the current user's home directory." };
        var force = new Option<bool>("--force") { Description = "Replace conflicting bundled files." };
        command.Options.Add(global);
        command.Options.Add(force);
        command.SetAction((parse, _) => JobCommands.Guard(() =>
        {
            string home = parse.GetValue(global)
                ? Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) : Directory.GetCurrentDirectory();
            if (string.IsNullOrEmpty(home))
                throw new DirectoryNotFoundException("User profile directory is unavailable.");
            string target = Path.Combine(home, ".agents", "skills", "statistica-cli");
            Install(target, parse.GetValue(force));
            Console.WriteLine(target);
            return Task.FromResult(0);
        }));
        root.Subcommands.Add(command);
    }

    internal static void Install(string target, bool force)
    {
        var files = new[] { "SKILL.md", "references/cli-guide.md" }.Select(resource =>
        {
            using var source = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource)
                ?? throw new InvalidOperationException($"Missing bundled resource: {resource}");
            using var bytes = new MemoryStream();
            source.CopyTo(bytes);
            string path = Path.Combine(target, resource);
            if (Directory.Exists(path))
                throw new IOException($"Skill file is a directory: {path}");
            byte[]? previous = File.Exists(path) ? File.ReadAllBytes(path) : null;
            byte[] content = bytes.ToArray();
            bool identical = previous is not null && content.AsSpan().SequenceEqual(previous);
            if (!identical && previous is not null && !force)
                throw new IOException($"Skill file conflicts with the bundled version: {path}. Use --force to replace it.");
            return (Path: path, Content: content, Previous: previous, Identical: identical);
        }).Where(file => !file.Identical).ToArray();

        var staged = new List<(string Temp, string Path, byte[]? Previous)>();
        var published = new Stack<(string Path, byte[]? Previous)>();
        try
        {
            foreach (var file in files)
            {
                string directory = Path.GetDirectoryName(file.Path)!;
                Directory.CreateDirectory(directory);
                string temp = Path.Combine(directory, $".statistica-{Guid.NewGuid():N}");
                using var stream = new FileStream(temp, FileMode.CreateNew);
                staged.Add((temp, file.Path, file.Previous));
                stream.Write(file.Content);
            }
            foreach (var file in staged)
            {
                File.Move(file.Temp, file.Path, overwrite: force);
                published.Push((file.Path, file.Previous));
            }
        }
        catch
        {
            foreach (var file in published)
            {
                if (file.Previous is null)
                    File.Delete(file.Path);
                else
                    File.WriteAllBytes(file.Path, file.Previous);
            }
            throw;
        }
        finally
        {
            foreach (var file in staged)
                File.Delete(file.Temp);
        }
    }
}
