using System.Runtime.InteropServices;

namespace StatisticaCli.Automation;

internal sealed record DataBlock(object?[,] Cells, int FirstRow = 0, int FirstColumn = 0);

internal enum OptionalArgument
{
    Missing
}

[StructLayout(LayoutKind.Explicit, Size = 24)]
internal unsafe struct Variant
{
    private const ushort Empty = 0, Null = 1, Int16 = 2, Int32 = 3, Float = 4, Double = 5,
        Date = 7, BStr = 8, Dispatch = 9, Error = 10, Boolean = 11, VariantType = 12,
        SByte = 16, Byte = 17, UInt16 = 18, UInt32 = 19, Int64 = 20, UInt64 = 21,
        Int = 22, UInt = 23, Array = 0x2000, ByRef = 0x4000;

    [FieldOffset(0)] private ushort type;
    [FieldOffset(8)] private nint pointer;
    [FieldOffset(8)] private long integer;
    [FieldOffset(8)] private float single;
    [FieldOffset(8)] private double number;

    internal static Variant From(object? value)
    {
        switch (value)
        {
            case null: return default;
            case int x: return new() { type = Int32, integer = x };
            case double x: return new() { type = Double, number = x };
            case bool x: return new() { type = Boolean, integer = x ? -1 : 0 };
            case OptionalArgument.Missing:
                const int parameterNotFound = unchecked((int)0x80020004);
                return new() { type = Error, integer = parameterNotFound };
            case string x:
                fixed (char* chars = x)
                {
                    nint text = Native.SysAllocStringLen(chars, (uint)x.Length);
                    return text != 0 ? new() { type = BStr, pointer = text } : throw new OutOfMemoryException();
                }
            case DispatchObject x:
                return new() { type = Dispatch, pointer = x.Retain() };
            case DataBlock x: return FromBlock(x);
            default: throw new NotSupportedException($"Unsupported Automation argument type: {value.GetType().Name}.");
        }
    }

    internal static object? Read(Variant* value, AutomationApartment apartment)
    {
        if ((value->type & ByRef) != 0)
        {
            Variant owned = default;
            try
            {
                AutomationException.ThrowIfFailed(Native.VariantCopyInd(&owned, value), "VariantCopyInd");
                return Read(&owned, apartment);
            }
            finally { Native.VariantClear(&owned); }
        }

        return value->type switch
        {
            Empty or Null => null,
            Int16 => (short)value->integer,
            Int32 or Int => (int)value->integer,
            Float => value->single,
            Double => value->number,
            Date => DateTime.FromOADate(value->number),
            BStr => ReadBStr(value->pointer) ?? "",
            Boolean => (short)value->integer != 0,
            SByte => (sbyte)value->integer,
            Byte => (byte)value->integer,
            UInt16 => (ushort)value->integer,
            UInt32 or UInt => (uint)value->integer,
            Int64 => value->integer,
            UInt64 => (ulong)value->integer,
            Dispatch => value->pointer == 0 ? null : DispatchObject.Retain(value->pointer, apartment),
            Array | VariantType => ReadBlock(value->pointer, apartment),
            _ => throw new NotSupportedException($"Unsupported Automation result type: 0x{value->type:X4}.")
        };
    }

    internal static string? ReadBStr(nint text) =>
        text == 0 ? null : new string((char*)text, 0, checked((int)Native.SysStringLen(text)));

    private static Variant FromBlock(DataBlock block)
    {
        int rows = block.Cells.GetLength(0), columns = block.Cells.GetLength(1);
        ArrayBound* bounds = stackalloc ArrayBound[2]
        {
            new((uint)rows, block.FirstRow),
            new((uint)columns, block.FirstColumn)
        };
        Variant result = new() { type = Array | VariantType, pointer = Native.SafeArrayCreate(VariantType, 2, bounds) };
        if (result.pointer == 0)
            throw new OutOfMemoryException();

        try
        {
            int* indices = stackalloc int[2];
            for (int column = 0; column < columns; column++)
            {
                indices[1] = checked(block.FirstColumn + column);
                for (int row = 0; row < rows; row++)
                {
                    indices[0] = checked(block.FirstRow + row);
                    Variant cell = From(block.Cells[row, column]);
                    try
                    {
                        AutomationException.ThrowIfFailed(
                            Native.SafeArrayPutElement(result.pointer, indices, &cell), "SafeArrayPutElement");
                    }
                    finally { Native.VariantClear(&cell); }
                }
            }
            return result;
        }
        catch
        {
            Native.VariantClear(&result);
            throw;
        }
    }

    private static DataBlock ReadBlock(nint array, AutomationApartment apartment)
    {
        if (array == 0 || Native.SafeArrayGetDim(array) != 2)
            throw new InvalidDataException("Expected a two-dimensional SAFEARRAY.");
        var (firstRow, rows) = GetBounds(array, 1);
        var (firstColumn, columns) = GetBounds(array, 2);
        var cells = new object?[rows, columns];
        AutomationException.ThrowIfFailed(Native.SafeArrayAccessData(array, out nint data), "SafeArrayAccessData");
        try
        {
            // Automation stores the leftmost dimension contiguously.
            var element = (Variant*)data;
            for (int column = 0; column < columns; column++)
                for (int row = 0; row < rows; row++, element++)
                {
                    if ((element->type & (Array | ByRef)) != 0 || element->type == Dispatch)
                        throw new InvalidDataException("Expected scalar cells in the SAFEARRAY.");
                    cells[row, column] = Read(element, apartment);
                }
        }
        finally { Native.SafeArrayUnaccessData(array); }
        return new(cells, firstRow, firstColumn);
    }

    private static (int Lower, int Count) GetBounds(nint array, uint dimension)
    {
        AutomationException.ThrowIfFailed(Native.SafeArrayGetLBound(array, dimension, out int lower), "SafeArrayGetLBound");
        AutomationException.ThrowIfFailed(Native.SafeArrayGetUBound(array, dimension, out int upper), "SafeArrayGetUBound");
        return (lower, checked(upper - lower + 1));
    }
}
