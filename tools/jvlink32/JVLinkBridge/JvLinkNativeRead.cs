// JvLinkNativeRead.cs – Pointer-buffer strategy for JVRead via IDispatch::Invoke.
//
// Allocates raw unmanaged buffers (via Marshal.AllocHGlobal) and passes them to
// JVRead as VT_BYREF|VT_UI1 variants (byte pointers), bypassing BSTR marshaling
// that previously caused SEHException or OutOfMemoryException.
//
// Encoding is controlled by JV_TEXT_ENCODING env var (default: 932 = Shift-JIS).

using System.Runtime.InteropServices;
using System.Text;

/// <summary>
/// Invokes JVRead via raw IDispatch::Invoke with unmanaged pointer buffers,
/// avoiding the SEHException / OutOfMemoryException from BSTR marshaling.
/// </summary>
static class JvLinkNativeReader
{
    // Buffer capacities
    const int DefaultCapacity  = 1 * 1024 * 1024; // 1 MB initial
    const int MaxCapacity      = 8 * 1024 * 1024; // 8 MB grow limit
    const int FilenameCapacity = 260;              // MAX_PATH

    // VARIANT type constants
    const ushort VT_I4    = 3;
    const ushort VT_UI1   = 17;
    const ushort VT_BYREF = 0x4000;

    // DISPATCH_METHOD flag for IDispatch::Invoke
    const ushort DISPATCH_METHOD = 1;

    // VARIANT is 16 bytes on x86: 2-byte vt, 6 bytes reserved, 8-byte data union.
    const int VariantSize    = 16;
    // DISPPARAMS is 4 x 4-byte fields = 16 bytes on x86.
    const int DispParamsSize = 16;

    // ── P/Invoke ─────────────────────────────────────────────────────────────

    [DllImport("kernel32.dll", EntryPoint = "RtlZeroMemory", SetLastError = false)]
    static extern void ZeroMemory(IntPtr dest, IntPtr size);

    // IDispatch vtable slot 5: GetIDsOfNames
    [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet = CharSet.Unicode)]
    delegate int GetIDsOfNamesDelegate(
        IntPtr pThis,
        ref Guid riid,
        IntPtr   rgszNames,   // LPOLESTR* (array of wide string pointers)
        uint     cNames,
        uint     lcid,
        IntPtr   rgDispId);   // DISPID* (out)

    // IDispatch vtable slot 6: Invoke
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    delegate int InvokeDelegate(
        IntPtr   pThis,
        int      dispIdMember,
        ref Guid riid,
        uint     lcid,
        ushort   wFlags,
        IntPtr   pDispParams,   // DISPPARAMS*
        IntPtr   pVarResult,    // VARIANT* (out, optional)
        IntPtr   pExcepInfo,    // EXCEPINFO* (out, optional)
        IntPtr   puArgErr);     // UINT* (out, optional)

    // ── Internal helpers ──────────────────────────────────────────────────────

    /// <summary>Gets the DISPID of a method by name via IDispatch::GetIDsOfNames.</summary>
    static int GetDispId(IntPtr pDisp, string methodName)
    {
        IntPtr vtable = Marshal.ReadIntPtr(pDisp);
        IntPtr fnPtr  = Marshal.ReadIntPtr(vtable, 5 * IntPtr.Size); // slot 5
        var    getIds = Marshal.GetDelegateForFunctionPointer<GetIDsOfNamesDelegate>(fnPtr);

        IntPtr pName   = Marshal.StringToHGlobalUni(methodName);
        IntPtr pNames  = Marshal.AllocHGlobal(IntPtr.Size); // array of 1 LPOLESTR
        IntPtr pDispId = Marshal.AllocHGlobal(4);
        try
        {
            Marshal.WriteIntPtr(pNames, pName);
            Marshal.WriteInt32(pDispId, 0);

            Guid riid = Guid.Empty;
            int  hr   = getIds(pDisp, ref riid, pNames, 1, 0x0409 /*en-US*/, pDispId);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);
            return Marshal.ReadInt32(pDispId);
        }
        finally
        {
            Marshal.FreeHGlobal(pName);
            Marshal.FreeHGlobal(pNames);
            Marshal.FreeHGlobal(pDispId);
        }
    }

    /// <summary>
    /// Calls IDispatch::Invoke for JVRead(buff, size, filename) with raw pointer VARIANTs.
    /// IDispatch passes args in reverse order: rgvarg[0]=filename, [1]=size, [2]=buff.
    /// Returns the integer return value of JVRead; size is read back by the caller from pSize.
    /// </summary>
    static int RawInvokeJVRead(
        IntPtr pDisp,
        int    dispId,
        IntPtr pBuff,
        IntPtr pSize,      // int* (in/out)
        IntPtr pFilename)
    {
        IntPtr vtable    = Marshal.ReadIntPtr(pDisp);
        IntPtr invokePtr = Marshal.ReadIntPtr(vtable, 6 * IntPtr.Size); // slot 6
        var    invoke    = Marshal.GetDelegateForFunctionPointer<InvokeDelegate>(invokePtr);

        IntPtr pVariants   = Marshal.AllocHGlobal(3 * VariantSize);
        IntPtr pVarResult  = Marshal.AllocHGlobal(VariantSize);
        IntPtr pDispParams = Marshal.AllocHGlobal(DispParamsSize);
        try
        {
            ZeroMemory(pVariants,   (IntPtr)(3 * VariantSize));
            ZeroMemory(pVarResult,  (IntPtr)VariantSize);
            ZeroMemory(pDispParams, (IntPtr)DispParamsSize);

            // rgvarg[0] = filename: VT_BYREF|VT_UI1 → byte* pointing at char buffer
            int o0 = 0 * VariantSize;
            Marshal.WriteInt16(pVariants, o0,     (short)(VT_BYREF | VT_UI1));
            Marshal.WriteIntPtr(pVariants + o0 + 8, pFilename);

            // rgvarg[1] = size: VT_BYREF|VT_I4 → int*
            int o1 = 1 * VariantSize;
            Marshal.WriteInt16(pVariants, o1,     (short)(VT_BYREF | VT_I4));
            Marshal.WriteIntPtr(pVariants + o1 + 8, pSize);

            // rgvarg[2] = buff: VT_BYREF|VT_UI1 → byte* pointing at char buffer
            int o2 = 2 * VariantSize;
            Marshal.WriteInt16(pVariants, o2,     (short)(VT_BYREF | VT_UI1));
            Marshal.WriteIntPtr(pVariants + o2 + 8, pBuff);

            // DISPPARAMS: { rgvarg, NULL, 3, 0 }
            Marshal.WriteIntPtr(pDispParams, 0,               pVariants);   // rgvarg
            Marshal.WriteIntPtr(pDispParams, IntPtr.Size,     IntPtr.Zero); // rgdispidNamedArgs
            Marshal.WriteInt32 (pDispParams, 2 * IntPtr.Size, 3);           // cArgs
            Marshal.WriteInt32 (pDispParams, 3 * IntPtr.Size, 0);           // cNamedArgs

            Guid riid = Guid.Empty;
            int  hr   = invoke(pDisp, dispId, ref riid, 0x0409, DISPATCH_METHOD,
                               pDispParams, pVarResult, IntPtr.Zero, IntPtr.Zero);
            if (hr < 0) Marshal.ThrowExceptionForHR(hr);

            // Extract integer return value from result VARIANT (JVRead returns VT_I4)
            short retVt = Marshal.ReadInt16(pVarResult, 0);
            return ((ushort)retVt == VT_I4)
                   ? Marshal.ReadInt32(pVarResult + 8)
                   : 0;
        }
        finally
        {
            Marshal.FreeHGlobal(pVariants);
            Marshal.FreeHGlobal(pVarResult);
            Marshal.FreeHGlobal(pDispParams);
        }
    }

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>
    /// Invokes JVRead on <paramref name="comObj"/> using raw unmanaged buffers via
    /// IDispatch::Invoke with VT_BYREF|VT_UI1 pointer arguments, avoiding BSTR marshaling.
    /// Returns (ret, size, buff_decoded, filename_decoded).
    /// </summary>
    public static (int ret, int size, string buff, string filename) JVRead(
        object         comObj,
        Action<string> log)
    {
        // Resolve text encoding from JV_TEXT_ENCODING env var (default: 932 = Shift-JIS)
        int codePage = 932;
        if (int.TryParse(Environment.GetEnvironmentVariable("JV_TEXT_ENCODING"),
                         out var cp) && cp > 0)
            codePage = cp;
        Encoding enc;
        try   { enc = Encoding.GetEncoding(codePage); }
        catch { enc = Encoding.GetEncoding(932); }

        int    capacity = DefaultCapacity;
        IntPtr pDisp    = Marshal.GetIDispatchForObject(comObj);
        try
        {
            int dispId = GetDispId(pDisp, "JVRead");
            log($"[NativeRead] enc={enc.EncodingName} capacity={capacity} dispId={dispId}");

            // Retry loop: grow buffer if the returned size exceeds current capacity.
            while (true)
            {
                IntPtr pBuff     = Marshal.AllocHGlobal(capacity + 1);
                IntPtr pFilename = Marshal.AllocHGlobal(FilenameCapacity + 1);
                IntPtr pSize     = Marshal.AllocHGlobal(4);
                bool   retry     = false;
                try
                {
                    ZeroMemory(pBuff,     (IntPtr)(capacity + 1));
                    ZeroMemory(pFilename, (IntPtr)(FilenameCapacity + 1));
                    Marshal.WriteInt32(pSize, 0);

                    int ret  = RawInvokeJVRead(pDisp, dispId, pBuff, pSize, pFilename);
                    int size = Marshal.ReadInt32(pSize);

                    log($"[NativeRead] ret={ret} size={size} capacity={capacity}");

                    if (size < 0)
                    {
                        log($"[NativeRead] negative size={size}; returning empty buff");
                        return (ret, 0, "", ReadAnsiString(pFilename, FilenameCapacity, enc));
                    }

                    if (size > capacity)
                    {
                        int newCap = Math.Min(size + 1, MaxCapacity);
                        if (newCap > capacity)
                        {
                            log($"[NativeRead] buffer too small ({capacity}); growing to {newCap}");
                            capacity = newCap;
                            retry    = true;
                        }
                        else
                        {
                            log($"[NativeRead] size={size} exceeds MaxCapacity={MaxCapacity}; truncating");
                        }
                    }

                    if (!retry)
                    {
                        int    actualSize  = Math.Min(size, capacity);
                        string buffStr     = "";
                        if (actualSize > 0)
                        {
                            var bytes      = new byte[actualSize];
                            Marshal.Copy(pBuff, bytes, 0, actualSize);
                            int previewLen = Math.Min(actualSize, 16);
                            log($"[NativeRead] first bytes: {BitConverter.ToString(bytes, 0, previewLen)}");
                            buffStr = enc.GetString(bytes);
                        }
                        string filenameStr = ReadAnsiString(pFilename, FilenameCapacity, enc);
                        return (ret, size, buffStr, filenameStr);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(pBuff);
                    Marshal.FreeHGlobal(pFilename);
                    Marshal.FreeHGlobal(pSize);
                }
                // retry == true: loop with increased capacity
            }
        }
        finally
        {
            Marshal.Release(pDisp);
        }
    }

    /// <summary>Reads an ANSI/Shift-JIS string from unmanaged memory up to the first NUL byte.</summary>
    static string ReadAnsiString(IntPtr ptr, int maxLen, Encoding enc)
    {
        if (ptr == IntPtr.Zero) return "";
        int len = 0;
        while (len < maxLen && Marshal.ReadByte(ptr, len) != 0) len++;
        if (len == 0) return "";
        var bytes = new byte[len];
        Marshal.Copy(ptr, bytes, 0, len);
        return enc.GetString(bytes);
    }
}
