using System.Runtime.InteropServices;

namespace StatisticaCli.Automation;

internal sealed class AutomationException(string operation, int code, string? description = null)
    : COMException($"{operation}: {description ?? Marshal.GetPInvokeErrorMessage(code)}", code)
{
    internal static void ThrowIfFailed(int code, string operation)
    {
        if (code < 0)
            throw new AutomationException(operation, code);
    }
}

internal sealed class AutomationApartment : IDisposable
{
    private readonly int thread = Environment.CurrentManagedThreadId;
    private bool disposed;

    internal Dictionary<(Guid Interface, string Member), int> DispatchIds { get; } = new();

    internal AutomationApartment()
    {
        const uint apartmentThreaded = 2;
        AutomationException.ThrowIfFailed(Native.CoInitializeEx(0, apartmentThreaded), "CoInitializeEx");
    }

    internal void CheckThread()
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        if (thread != Environment.CurrentManagedThreadId)
            throw new InvalidOperationException("Automation objects must be used on their owning STA thread.");
    }

    public void Dispose()
    {
        if (disposed)
            return;
        CheckThread();
        Native.CoUninitialize();
        disposed = true;
    }
}

internal sealed unsafe class DispatchObject : IDisposable
{
    private const uint UserDefaultLocale = 0x400;
    private const ushort Method = 1, PropertyGet = 2, PropertyPut = 4;
    private readonly AutomationApartment apartment;
    private nint instance;
    private Guid? interfaceId;

    private DispatchObject(nint instance, AutomationApartment apartment)
    {
        this.instance = instance;
        this.apartment = apartment;
    }

    internal static DispatchObject Activate(AutomationApartment apartment, string progId, bool localServer)
    {
        apartment.CheckThread();
        const uint inProcessServer = 1, outOfProcessServer = 4;
        var dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
        AutomationException.ThrowIfFailed(Native.CLSIDFromProgID(progId, out Guid classId), $"CLSIDFromProgID({progId})");
        AutomationException.ThrowIfFailed(Native.CoCreateInstance(classId, 0,
            localServer ? outOfProcessServer : inProcessServer, dispatchId, out nint result), $"CoCreateInstance({progId})");
        return new(result, apartment);
    }

    internal T Get<T>(string member, params object?[] indices) =>
        ConvertResult<T>(Invoke(member, PropertyGet, indices, readResult: true), member);

    internal object? GetValue(string member, params object?[] indices) =>
        Invoke(member, PropertyGet, indices, readResult: true);

    internal T Call<T>(string member, params object?[] args) =>
        ConvertResult<T>(Invoke(member, Method, args, readResult: true), member);

    internal void Call(string member, params object?[] args) => Invoke(member, Method, args);

    internal void Set(string member, object? value, params object?[] indices)
    {
        var args = new object?[indices.Length + 1];
        indices.CopyTo(args, 0);
        args[^1] = value;
        Invoke(member, PropertyPut, args);
    }

    private static T ConvertResult<T>(object? value, string member)
    {
        if (value is T expected)
            return expected;
        if (value is IConvertible && typeof(T) == typeof(int))
            return (T)(object)Convert.ToInt32(value);
        if (value is IConvertible && typeof(T) == typeof(double))
            return (T)(object)Convert.ToDouble(value);
        (value as DispatchObject)?.Dispose();
        throw new InvalidDataException($"Unexpected result type for {member}. Expected {typeof(T).Name}.");
    }

    private object? Invoke(string member, ushort kind, object?[] values, bool readResult = false)
    {
        apartment.CheckThread();
        ObjectDisposedException.ThrowIf(instance == 0, this);
        Native.PumpMessages();
        int id = FindDispatchId(member);
        Span<Variant> arguments = values.Length <= 16 ? stackalloc Variant[values.Length] : new Variant[values.Length];
        arguments.Clear();
        Variant result = default;
        ExceptionInfo error = default;
        try
        {
            for (int i = 0; i < values.Length; i++)
                arguments[values.Length - 1 - i] = Variant.From(values[i]);

            fixed (Variant* args = arguments)
            {
                const int propertyPutId = -3;
                int named = propertyPutId;
                var parameters = new DispatchParameters
                {
                    Arguments = args,
                    Count = (uint)values.Length,
                    NamedIds = kind == PropertyPut ? &named : null,
                    NamedCount = kind == PropertyPut ? 1u : 0
                };
                Guid iid = Guid.Empty;
                uint argError = uint.MaxValue;
                var invoke = (delegate* unmanaged[Stdcall]<nint, int, Guid*, uint, ushort,
                    DispatchParameters*, Variant*, ExceptionInfo*, uint*, int>)Slot(instance, 6);
                int code = invoke(instance, id, &iid, UserDefaultLocale, kind,
                    &parameters, kind == PropertyPut ? null : &result, &error, &argError);
                if (code < 0)
                {
                    if (error.Fill != null)
                        error.Fill(&error);
                    string? description = Variant.ReadBStr(error.Description);
                    string? source = Variant.ReadBStr(error.Source);
                    const int typeMismatch = unchecked((int)0x80020005), parameterNotFound = unchecked((int)0x80020004);
                    string operation = code is (typeMismatch or parameterNotFound) && argError < values.Length
                        ? $"IDispatch.Invoke({member}, argument {values.Length - argError})"
                        : $"IDispatch.Invoke({member})";
                    if (error.Status != 0)
                        description = $"[0x{error.Status:X8}] {description ?? Marshal.GetPInvokeErrorMessage(error.Status)}";
                    else if (error.Code != 0)
                        description = $"[{error.Code}] {description}";
                    if (!string.IsNullOrEmpty(source))
                        description = description is null ? source : $"{source}: {description}";
                    throw new AutomationException(operation, code, description);
                }
            }
            return readResult ? Variant.Read(&result, apartment) : null;
        }
        finally
        {
            Native.VariantClear(&result);
            Native.SysFreeString(error.Source);
            Native.SysFreeString(error.Description);
            Native.SysFreeString(error.HelpFile);
            fixed (Variant* args = arguments)
                for (int i = 0; i < arguments.Length; i++)
                    Native.VariantClear(&args[i]);
        }
    }

    private int FindDispatchId(string member)
    {
        Guid identity = interfaceId ??= GetInterfaceId();
        if (identity != Guid.Empty && apartment.DispatchIds.TryGetValue((identity, member), out int cached))
            return cached;

        fixed (char* name = member)
        {
            char* names = name;
            Guid iid = Guid.Empty;
            int id;
            var getId = (delegate* unmanaged[Stdcall]<nint, Guid*, char**, uint, uint, int*, int>)Slot(instance, 5);
            AutomationException.ThrowIfFailed(getId(instance, &iid, &names, 1, UserDefaultLocale, &id),
                $"IDispatch.GetIDsOfNames({member})");
            if (identity != Guid.Empty)
                apartment.DispatchIds[(identity, member)] = id;
            return id;
        }
    }

    private Guid GetInterfaceId()
    {
        nint typeInfo = 0;
        var getTypeInfo = (delegate* unmanaged[Stdcall]<nint, uint, uint, nint*, int>)Slot(instance, 4);
        if (getTypeInfo(instance, 0, UserDefaultLocale, &typeInfo) < 0 || typeInfo == 0)
            return Guid.Empty;
        try
        {
            Guid* attributes = null;
            var getAttributes = (delegate* unmanaged[Stdcall]<nint, Guid**, int>)Slot(typeInfo, 3);
            if (getAttributes(typeInfo, &attributes) < 0 || attributes == null)
                return Guid.Empty;
            try { return *attributes; } // TYPEATTR starts with the interface GUID.
            finally
            {
                var releaseAttributes = (delegate* unmanaged[Stdcall]<nint, Guid*, void>)Slot(typeInfo, 19);
                releaseAttributes(typeInfo, attributes);
            }
        }
        finally { Release(typeInfo); }
    }

    internal nint Retain()
    {
        apartment.CheckThread();
        ObjectDisposedException.ThrowIf(instance == 0, this);
        AddRef(instance);
        return instance;
    }

    internal static DispatchObject Retain(nint instance, AutomationApartment apartment)
    {
        AddRef(instance);
        return new(instance, apartment);
    }

    public void Dispose()
    {
        if (instance == 0)
            return;
        apartment.CheckThread();
        Release(instance);
        instance = 0;
    }

    private static nint Slot(nint instance, int index) => (*(nint**)instance)[index];
    private static void AddRef(nint instance) => ((delegate* unmanaged[Stdcall]<nint, uint>)Slot(instance, 1))(instance);
    private static void Release(nint instance) => ((delegate* unmanaged[Stdcall]<nint, uint>)Slot(instance, 2))(instance);

    [StructLayout(LayoutKind.Sequential)]
    private struct DispatchParameters
    {
        internal Variant* Arguments;
        internal int* NamedIds;
        internal uint Count;
        internal uint NamedCount;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct ExceptionInfo
    {
        internal ushort Code;
        internal ushort Reserved;
        internal nint Source;
        internal nint Description;
        internal nint HelpFile;
        internal uint HelpContext;
        internal nint ReservedPointer;
        internal delegate* unmanaged[Stdcall]<ExceptionInfo*, int> Fill;
        internal int Status;
    }
}
