using System.Runtime.InteropServices;

namespace StatisticaCli.Automation;

internal static unsafe partial class Native
{
    [LibraryImport("ole32.dll")]
    internal static partial int CoInitializeEx(nint reserved, uint flags);

    [LibraryImport("ole32.dll")]
    internal static partial void CoUninitialize();

    [LibraryImport("ole32.dll", StringMarshalling = StringMarshalling.Utf16)]
    internal static partial int CLSIDFromProgID(string progId, out Guid classId);

    [LibraryImport("ole32.dll")]
    internal static partial int CoCreateInstance(in Guid classId, nint outer, uint context,
        in Guid interfaceId, out nint result);

    [LibraryImport("oleaut32.dll")]
    internal static partial nint SysAllocStringLen(char* text, uint length);

    [LibraryImport("oleaut32.dll")]
    internal static partial uint SysStringLen(nint text);

    [LibraryImport("oleaut32.dll")]
    internal static partial void SysFreeString(nint text);

    [LibraryImport("oleaut32.dll")]
    internal static partial int VariantClear(Variant* value);

    [LibraryImport("oleaut32.dll")]
    internal static partial int VariantCopyInd(Variant* destination, Variant* source);

    [LibraryImport("oleaut32.dll")]
    internal static partial nint SafeArrayCreate(ushort type, uint dimensions, ArrayBound* bounds);

    [LibraryImport("oleaut32.dll")]
    internal static partial uint SafeArrayGetDim(nint array);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayGetLBound(nint array, uint dimension, out int lower);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayGetUBound(nint array, uint dimension, out int upper);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayPutElement(nint array, int* indices, Variant* value);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayAccessData(nint array, out nint data);

    [LibraryImport("oleaut32.dll")]
    internal static partial int SafeArrayUnaccessData(nint array);

    [LibraryImport("user32.dll")]
    internal static partial int PeekMessageW(out WindowMessage message, nint window, uint min, uint max, uint flags);

    [LibraryImport("user32.dll")]
    internal static partial int TranslateMessage(in WindowMessage message);

    [LibraryImport("user32.dll")]
    internal static partial nint DispatchMessageW(in WindowMessage message);

    internal static void PumpMessages()
    {
        const uint remove = 1;
        while (PeekMessageW(out var message, 0, 0, 0, remove) != 0)
        {
            TranslateMessage(message);
            DispatchMessageW(message);
        }
    }
}

[StructLayout(LayoutKind.Sequential)]
internal struct WindowMessage
{
    private nint window;
    private uint message;
    private nuint wParam;
    private nint lParam;
    private uint time;
    private int x;
    private int y;
    private uint privateData;
}

[StructLayout(LayoutKind.Sequential)]
internal readonly record struct ArrayBound(uint Length, int Lower);
