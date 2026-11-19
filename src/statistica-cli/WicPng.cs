using StatisticaCli.Automation;

namespace StatisticaCli;

internal static unsafe class WicPng
{
    private static readonly Guid ImagingFactory = new("CACAF262-9370-4615-A13B-9F5539DA4C0A");
    private static readonly Guid ImagingFactoryInterface = new("EC5EC8A9-C395-4314-9C77-54D7A935FF70");
    private static readonly Guid ContainerPng = new("1B7CFAF4-713F-473C-BBCD-6137425FAEAF");
    private static readonly Guid Bgr24 = new("6FDDC324-4E03-4BFE-B185-3D77768DC90C");

    internal static void Save(string path, nint pixels, int bitmapWidth, int left, int top, int width, int height)
    {
        nint factory = 0, stream = 0, encoder = 0, frame = 0, properties = 0;
        try
        {
            Check(Native.CoCreateInstance(ImagingFactory, 0, 1, ImagingFactoryInterface, out factory), "CoCreateInstance(WIC)");
            Check(((delegate* unmanaged[Stdcall]<nint, nint*, int>)Slot(factory, 14))(factory, &stream), "CreateStream");
            fixed (char* filename = path)
                Check(((delegate* unmanaged[Stdcall]<nint, char*, uint, int>)Slot(stream, 15))(stream, filename, 0x40000000),
                    "InitializeFromFilename");

            Guid container = ContainerPng;
            Check(((delegate* unmanaged[Stdcall]<nint, Guid*, Guid*, nint*, int>)Slot(factory, 8))
                (factory, &container, null, &encoder), "CreateEncoder");
            const int noCache = 2;
            Check(((delegate* unmanaged[Stdcall]<nint, nint, int, int>)Slot(encoder, 3))(encoder, stream, noCache), "Encoder.Initialize");
            Check(((delegate* unmanaged[Stdcall]<nint, nint*, nint*, int>)Slot(encoder, 10))
                (encoder, &frame, &properties), "CreateNewFrame");
            Check(((delegate* unmanaged[Stdcall]<nint, nint, int>)Slot(frame, 3))(frame, properties), "Frame.Initialize");
            Check(((delegate* unmanaged[Stdcall]<nint, uint, uint, int>)Slot(frame, 4))(frame, (uint)width, (uint)height), "SetSize");
            Guid format = Bgr24;
            Check(((delegate* unmanaged[Stdcall]<nint, Guid*, int>)Slot(frame, 6))(frame, &format), "SetPixelFormat");
            if (format != Bgr24)
                throw new IOException("WIC did not accept the 24-bit BGR pixel format.");

            int stride = checked(width * 3);
            var buffer = new byte[stride];
            fixed (byte* row = buffer)
                for (int y = 0; y < height; y++)
                {
                    var source = (byte*)pixels + checked(((y + top) * bitmapWidth + left) * 4);
                    for (int x = 0; x < width; x++, source += 4)
                    {
                        row[x * 3] = source[0];
                        row[x * 3 + 1] = source[1];
                        row[x * 3 + 2] = source[2];
                    }
                    Check(((delegate* unmanaged[Stdcall]<nint, uint, uint, uint, byte*, int>)Slot(frame, 10))
                        (frame, 1, (uint)stride, (uint)stride, row), "WritePixels");
                }
            Check(((delegate* unmanaged[Stdcall]<nint, int>)Slot(frame, 12))(frame), "Frame.Commit");
            Check(((delegate* unmanaged[Stdcall]<nint, int>)Slot(encoder, 11))(encoder), "Encoder.Commit");
        }
        finally
        {
            Release(properties);
            Release(frame);
            Release(encoder);
            Release(stream);
            Release(factory);
        }
    }

    private static void Check(int result, string operation) => AutomationException.ThrowIfFailed(result, $"WIC.{operation}");
    private static nint Slot(nint value, int index) => (*(nint**)value)[index];
    private static void Release(nint value)
    {
        if (value != 0)
            ((delegate* unmanaged[Stdcall]<nint, uint>)Slot(value, 2))(value);
    }
}
