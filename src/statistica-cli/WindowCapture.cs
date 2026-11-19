using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using StatisticaCli.Automation;

namespace StatisticaCli;

internal static unsafe partial class WindowCapture
{
    [StructLayout(LayoutKind.Sequential)]
    private struct Rect
    {
        internal int Left, Top, Right, Bottom;
        internal readonly int Width => Right - Left;
        internal readonly int Height => Bottom - Top;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct BitmapHeader
    {
        internal uint Size;
        internal int Width, Height;
        internal ushort Planes, Bits;
        internal uint Compression, ImageSize;
        internal int XPelsPerMeter, YPelsPerMeter;
        internal uint ColorsUsed, ColorsImportant;
    }

    internal static void Save(DispatchObject app, int pid, JobOptions options, string input, string stage,
        Action<string, string, Action> call)
    {
        if (!HasInteractiveDesktop())
            throw new InvalidOperationException("Screenshot requires an unlocked interactive Windows desktop.");
        var (collection, activeProperty) = Path.GetExtension(input).ToLowerInvariant() switch
        {
            ".sta" => ("Spreadsheets", "ActiveSpreadsheet"),
            ".stg" => ("Graphs", "ActiveGraph"),
            ".stw" => ("Workbooks", "ActiveWorkbook"),
            _ => throw new InvalidDataException("Unsupported document type for screenshot.")
        };
        using var document = Documents.Open(app, collection, input);
        document.LeaveForQuit = true;
        app.Set("Visible", true);
        document.Value.Set("Visible", true);
        document.Value.Call("Activate");
        using var window = document.Value.Get<DispatchObject>("Window");
        nint mainHandle = (nint)app.Get<long>("Handle");
        nint docHandle = (nint)window.Get<long>("Handle");
        var clock = Stopwatch.StartNew();
        BitmapBuffer? bitmap = null;
        var previous = new int[72 * 54];
        var current = new int[previous.Length];
        bool hasPrevious = false;
        Rect previousRegion = default;
        try
        {
            do
            {
                Native.PumpMessages();
                bool stableCandidate = false;
                if (Bounds(mainHandle, pid, out var main) && Bounds(docHandle, pid, out var target))
                {
                    var region = new Rect
                    {
                        Left = Math.Max(0, target.Left - main.Left),
                        Top = Math.Max(0, target.Top - main.Top),
                        Right = Math.Min(main.Width, target.Right - main.Left),
                        Bottom = Math.Min(main.Height, target.Bottom - main.Top)
                    };
                    if (region.Width >= 160 && region.Height >= 100 && IsActive(app, activeProperty, docHandle))
                    {
                        if (bitmap is null || bitmap.Width != main.Width || bitmap.Height != main.Height)
                        {
                            var resized = new BitmapBuffer(main.Width, main.Height);
                            bitmap?.Dispose();
                            bitmap = resized;
                            hasPrevious = false;
                        }
                        new Span<int>((void*)bitmap.Pixels, checked(main.Width * main.Height)).Clear();
                        bool copied = false;
                        call("Capture", "PrintWindow", () => copied = PrintWindow(mainHandle, bitmap.Dc, 0) != 0);
                        stableCandidate = copied && Sample(bitmap, region, current);
                        if (stableCandidate && hasPrevious && region.Left == previousRegion.Left &&
                            region.Top == previousRegion.Top && region.Right == previousRegion.Right && region.Bottom == previousRegion.Bottom &&
                            current.AsSpan().SequenceEqual(previous) && HasInteractiveDesktop() && IsActive(app, activeProperty, docHandle))
                        {
                            var crop = options.CropToDocument ? region : new Rect { Right = main.Width, Bottom = main.Height };
                            call("Capture", "WIC PNG", () =>
                                WicPng.Save(stage, bitmap.Pixels, bitmap.Width, crop.Left, crop.Top, crop.Width, crop.Height));
                            return;
                        }
                        previousRegion = region;
                        (previous, current) = (current, previous);
                    }
                }
                hasPrevious = stableCandidate;
                Thread.Sleep(160);
            } while (clock.Elapsed.TotalSeconds < options.CaptureTimeoutSeconds);
            throw new IOException("The active document did not produce a stable, nonblank frame.");
        }
        finally { bitmap?.Dispose(); }
    }

    private static bool Bounds(nint handle, int pid, out Rect bounds)
    {
        bounds = default;
        return GetWindowThreadProcessId(handle, out uint owner) != 0 && owner == pid && IsWindowVisible(handle) != 0 &&
            IsIconic(handle) == 0 && GetWindowRect(handle, out bounds) != 0;
    }

    private static bool IsActive(DispatchObject app, string activeProperty, nint handle)
    {
        using var active = (DispatchObject?)app.GetValue(activeProperty);
        using var window = (DispatchObject?)active?.GetValue("Window");
        return window is not null && (nint)window.Get<long>("Handle") == handle;
    }

    private static bool Sample(BitmapBuffer bitmap, Rect region, int[] samples)
    {
        var colors = new HashSet<int>();
        var pixels = (int*)bitmap.Pixels;
        int left = region.Left + 12, top = region.Top + 32;
        int width = region.Width - 24, height = region.Height - 44;
        int first = pixels[top * bitmap.Width + left] & 0xFFFFFF;
        int different = 0;
        for (int y = 0; y < 54; y++)
            for (int x = 0; x < 72; x++)
            {
                int pixel = pixels[(top + y * height / 54) * bitmap.Width + left + x * width / 72] & 0xFFFFFF;
                samples[y * 72 + x] = pixel;
                if (colors.Count < 3)
                    colors.Add(pixel);
                if (pixel != first)
                    different++;
            }
        return colors.Count >= 3 && different >= 8;
    }

    internal static bool HasInteractiveDesktop()
    {
        if (!Environment.UserInteractive)
            return false;
        const uint readObjects = 1, objectName = 2;
        nint input = OpenInputDesktop(0, 0, readObjects);
        if (input == 0)
            return false;
        try
        {
            char* name = stackalloc char[256];
            return GetUserObjectInformation(input, objectName, name, 512, out uint bytes) != 0 && bytes >= 2 &&
                new ReadOnlySpan<char>(name, (int)bytes / 2 - 1).Equals("Default", StringComparison.OrdinalIgnoreCase);
        }
        finally { CloseDesktop(input); }
    }

    private sealed class BitmapBuffer : IDisposable
    {
        internal int Width { get; }
        internal int Height { get; }
        internal nint Dc { get; }
        internal nint Pixels { get; }
        private readonly nint bitmap, previous;

        internal BitmapBuffer(int width, int height)
        {
            Width = width;
            Height = height;
            Dc = CreateCompatibleDC(0);
            if (Dc == 0)
                throw new Win32Exception();
            try
            {
                var header = new BitmapHeader { Size = (uint)sizeof(BitmapHeader), Width = width, Height = -height, Planes = 1, Bits = 32 };
                bitmap = CreateDIBSection(Dc, &header, 0, out nint pixels, 0, 0);
                if (bitmap == 0 || pixels == 0)
                    throw new Win32Exception();
                Pixels = pixels;
                previous = SelectObject(Dc, bitmap);
                if (previous is 0 or -1)
                    throw new Win32Exception();
            }
            catch
            {
                if (bitmap != 0)
                    DeleteObject(bitmap);
                DeleteDC(Dc);
                throw;
            }
        }

        public void Dispose()
        {
            SelectObject(Dc, previous);
            DeleteObject(bitmap);
            DeleteDC(Dc);
        }
    }

    [LibraryImport("user32.dll")] private static partial int GetWindowRect(nint window, out Rect bounds);
    [LibraryImport("user32.dll")] private static partial int IsWindowVisible(nint window);
    [LibraryImport("user32.dll")] private static partial int IsIconic(nint window);
    [LibraryImport("user32.dll")] private static partial uint GetWindowThreadProcessId(nint window, out uint pid);
    [LibraryImport("user32.dll")] private static partial int PrintWindow(nint window, nint dc, uint flags);
    [LibraryImport("user32.dll")] private static partial nint OpenInputDesktop(uint flags, int inherit, uint access);
    [LibraryImport("user32.dll")] private static partial int CloseDesktop(nint desktop);
    [LibraryImport("user32.dll", EntryPoint = "GetUserObjectInformationW")]
    private static partial int GetUserObjectInformation(nint handle, uint index, char* info, uint size, out uint required);
    [LibraryImport("gdi32.dll", SetLastError = true)] private static partial nint CreateCompatibleDC(nint dc);
    [LibraryImport("gdi32.dll")] private static partial int DeleteDC(nint dc);
    [LibraryImport("gdi32.dll", SetLastError = true)]
    private static partial nint CreateDIBSection(nint dc, BitmapHeader* info, uint usage, out nint bits, nint section, uint offset);
    [LibraryImport("gdi32.dll", SetLastError = true)] private static partial nint SelectObject(nint dc, nint handle);
    [LibraryImport("gdi32.dll")] private static partial int DeleteObject(nint handle);
}
