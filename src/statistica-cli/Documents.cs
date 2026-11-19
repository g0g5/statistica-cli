using System.Buffers.Binary;
using System.Xml;
using StatisticaCli.Automation;
using System.Runtime.InteropServices;

namespace StatisticaCli;

internal sealed class DocumentCleanupException(string operation, COMException error)
    : COMException($"{operation}: {error.Message}", error.HResult);

internal sealed class Documents : IDisposable
{
    internal DispatchObject Value { get; }
    internal bool LeaveForQuit { get; set; }

    private Documents(DispatchObject value) => Value = value;

    internal static Documents Open(DispatchObject app, string collection, string path, bool writable = false)
    {
        using var items = app.Get<DispatchObject>(collection);
        return new(items.Call<DispatchObject>(writable ? "Open" : "OpenReadOnly", path, false));
    }

    internal static Documents New(DispatchObject app, string collection)
    {
        using var items = app.Get<DispatchObject>(collection);
        return new(items.Call<DispatchObject>("New"));
    }

    internal static Documents Own(DispatchObject value) => new(value);

    internal static void Copy(string source, string stage)
    {
        File.Copy(source, stage);
        File.SetAttributes(stage, FileAttributes.Normal);
    }

    internal void SaveAs(string path)
    {
        if (!Value.Call<bool>("SaveAs", path, true))
            throw new IOException($"SaveAs returned false for {path}.");
    }

    internal void RequireWritable()
    {
        if (Value.Get<bool>("ReadOnly"))
            throw new IOException($"The working spreadsheet is read-only: {Value.Get<string>("FullName")}");
    }

    internal void Save()
    {
        RequireWritable();
        Value.Call("Save");
        if (!Value.Get<bool>("Saved"))
            throw new IOException("STATISTICA still reports unsaved changes after Save().");
    }

    internal static void VerifyOutput(string path)
    {
        using var stream = File.OpenRead(path);
        if (stream.Length == 0)
            throw new IOException($"Expected a nonempty output file: {path}");
        switch (Path.GetExtension(path).ToLowerInvariant())
        {
            case ".png":
                Span<byte> header = stackalloc byte[24];
                stream.ReadExactly(header);
                if (!header[..8].SequenceEqual(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }) ||
                    !header[12..16].SequenceEqual("IHDR"u8) ||
                    BinaryPrimitives.ReadInt32BigEndian(header[16..20]) <= 0 ||
                    BinaryPrimitives.ReadInt32BigEndian(header[20..24]) <= 0)
                    throw new IOException("Invalid PNG header or image dimensions.");
                break;
            case ".svg":
                using (var xml = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore }))
                {
                    xml.MoveToContent();
                    if (xml.LocalName != "svg" || xml.NamespaceURI != "http://www.w3.org/2000/svg")
                        throw new IOException("The output is missing its SVG root element.");
                    while (xml.Read()) { }
                }
                break;
        }
    }

    public void Dispose()
    {
        using (Value)
        {
            try
            {
                if (LeaveForQuit)
                    Value.Set("Saved", true);
                else
                    Value.Call("Close", false);
            }
            catch (COMException error)
            {
                throw new DocumentCleanupException(LeaveForQuit ? "Document.Saved" : "Document.Close", error);
            }
        }
    }
}
