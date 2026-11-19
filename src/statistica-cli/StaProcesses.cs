using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace StatisticaCli;

internal static unsafe partial class StaProcesses
{
    internal static int[] Baseline()
    {
        var processes = Process.GetProcessesByName("statist");
        try { return Array.ConvertAll(processes, process => process.Id); }
        finally
        {
            foreach (var process in processes)
                process.Dispose();
        }
    }

    internal static DesktopProcess[] Snapshot()
    {
        var entries = new List<DesktopProcess>();
        var processes = Process.GetProcessesByName("statist");
        foreach (var process in processes)
        {
            using (process)
            {
                try
                {
                    entries.Add(new(process.Id, process.SessionId, process.StartTime.ToUniversalTime(),
                        ReadCommandLine(process.SafeHandle)));
                }
                catch (Exception e) when (e is Win32Exception or InvalidOperationException) { }
            }
        }
        return entries.ToArray();
    }

    internal static ServerIdentity Identify(int pid)
    {
        using var process = Process.GetProcessById(pid);
        return new(pid, process.SessionId, process.StartTime.ToUniversalTime().Ticks);
    }

    internal static bool StopOwned(ServerIdentity identity, int[] baseline, long workerStarted, int session)
    {
        if (identity.Session != session || baseline.Contains(identity.Pid) || identity.StartTicks < workerStarted)
            return false;

        try
        {
            using var process = Process.GetProcessById(identity.Pid);
            // Keep the process handle open throughout verification and termination.
            SafeProcessHandle handle = process.SafeHandle;
            if (process.ProcessName != "statist" || process.SessionId != identity.Session ||
                process.StartTime.ToUniversalTime().Ticks != identity.StartTicks ||
                !IsEmbeddingServer(ReadCommandLine(handle)))
                return false;

            if (!TerminateProcess(handle, 1))
                throw new Win32Exception();
            return process.WaitForExit(3000);
        }
        catch (ArgumentException) { return false; } // The server has already exited.
        catch (InvalidOperationException) { return false; }
        catch (Win32Exception error)
        {
            Console.Error.WriteLine($"Could not terminate verified STATISTICA PID {identity.Pid}: {error.Message}");
            return false;
        }
    }

    private static bool IsEmbeddingServer(string? commandLine)
    {
        if (string.IsNullOrWhiteSpace(commandLine))
            return false;
        nint argv = CommandLineToArgvW(commandLine, out int argc);
        if (argv == 0)
            return false;
        try
        {
            var args = (nint*)argv;
            return argc == 2 &&
                string.Equals(Path.GetFileName(Marshal.PtrToStringUni(args[0])), "statist.exe", StringComparison.OrdinalIgnoreCase) &&
                string.Equals(Marshal.PtrToStringUni(args[1]), "-Embedding", StringComparison.OrdinalIgnoreCase);
        }
        finally { LocalFree(argv); }
    }

    private static string? ReadCommandLine(SafeProcessHandle process)
    {
        const int processCommandLineInformation = 60;
        NtQueryInformationProcess(process, processCommandLineInformation, null, 0, out int length);
        if (length < sizeof(UnicodeString) || length > 128 * 1024)
            return null;

        var buffer = new byte[length];
        fixed (byte* start = buffer)
        {
            if (NtQueryInformationProcess(process, processCommandLineInformation, start, length, out _) < 0)
                return null;
            var text = (UnicodeString*)start;
            if (text->Length % 2 != 0 || text->Buffer < start ||
                text->Buffer > start + length - text->Length)
                return null;
            return new string((char*)text->Buffer, 0, text->Length / 2);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct UnicodeString
    {
        internal ushort Length;
        internal ushort MaximumLength;
        internal byte* Buffer;
    }

    [LibraryImport("ntdll.dll")]
    private static partial int NtQueryInformationProcess(SafeProcessHandle process, int kind, void* buffer,
        int length, out int returned);

    [LibraryImport("shell32.dll", StringMarshalling = StringMarshalling.Utf16)]
    private static partial nint CommandLineToArgvW(string commandLine, out int count);

    [LibraryImport("kernel32.dll")]
    private static partial nint LocalFree(nint memory);

    [LibraryImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static partial bool TerminateProcess(SafeProcessHandle process, uint exitCode);
}
