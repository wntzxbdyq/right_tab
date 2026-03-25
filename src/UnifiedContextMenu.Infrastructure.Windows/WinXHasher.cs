using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Text;

namespace UnifiedContextMenu.Infrastructure.Windows;

internal static class WinXHasher
{
    [DllImport("shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError = true)]
    private static extern int HashData(
        [In][MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 1)] byte[] pbData, int cbData,
        [Out][MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 3)] byte[] pbHash, int cbHash);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern uint SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IBindCtx? pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IShellItem ppv);

    [DllImport("propsys.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int PSGetPropertyKeyFromName([In][MarshalAs(UnmanagedType.LPWStr)] string pszCanonicalName, out PropertyKey propkey);

    [ComImport]
    [Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IPropertyStore
    {
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetCount([Out] out uint cProps);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetAt([In] uint iProp, out PropertyKey pkey);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void GetValue([In] ref PropertyKey key, out PropVariant pv);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void SetValue([In] ref PropertyKey key, [In] ref PropVariant pv);
        [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
        void Commit();
    }

    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    private interface IShellItem
    {
        void BindToHandler(IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid bhid, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, out IntPtr ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    [ComImport]
    [SuppressUnmanagedCodeSecurity]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("7e9fb0d3-919f-4307-ab2e-9b1860310c93")]
    private interface IShellItem2 : IShellItem
    {
        [return: MarshalAs(UnmanagedType.Interface)] object BindToHandler(IBindCtx pbc, [In] ref Guid bhid, [In] ref Guid riid);
        new IShellItem GetParent();
        [return: MarshalAs(UnmanagedType.LPWStr)] new string GetDisplayName(uint sigdnName);
        new uint GetAttributes(uint sfgaoMask);
        new int Compare(IShellItem psi, uint hint);
        [return: MarshalAs(UnmanagedType.Interface)] IPropertyStore GetPropertyStore(GPS flags, [In] ref Guid riid);
        [return: MarshalAs(UnmanagedType.Interface)] object GetPropertyStoreWithCreateObject(GPS flags, [MarshalAs(UnmanagedType.IUnknown)] object punkCreateObject, [In] ref Guid riid);
        [return: MarshalAs(UnmanagedType.Interface)] object GetPropertyStoreForKeys(IntPtr rgKeys, uint cKeys, GPS flags, [In] ref Guid riid);
        [return: MarshalAs(UnmanagedType.Interface)] object GetPropertyDescriptionList(IntPtr keyType, [In] ref Guid riid);
        void Update(IBindCtx pbc);
        [SecurityCritical] void GetProperty(IntPtr key, [In][Out] PropVariant pv);
        Guid GetCLSID(IntPtr key);
        System.Runtime.InteropServices.ComTypes.FILETIME GetFileTime(IntPtr key);
        int GetInt32(IntPtr key);
        [return: MarshalAs(UnmanagedType.LPWStr)] string GetString(PropertyKey key);
        uint GetUInt32(IntPtr key);
        ulong GetUInt64(IntPtr key);
        [return: MarshalAs(UnmanagedType.Bool)] bool GetBool(IntPtr key);
    }

    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    private struct PropertyKey { public Guid GUID; public int PID; }

    [StructLayout(LayoutKind.Explicit, Pack = 1)]
    private struct PropVariant
    {
        [FieldOffset(0)] public VarEnum VarType;
        [FieldOffset(8)] public uint ulVal;
    }

    [Flags]
    private enum GPS { READWRITE = 0x00000002 }

    private static readonly Dictionary<string, string> GeneralizePathDic = new()
    {
        { "%ProgramFiles%", "{905e63b6-c1bf-494e-b29c-65b732d3d21a}" },
        { "%SystemRoot%\\System32", "{1ac14e77-02e7-4e5d-b744-2eb1ae5198b7}" },
        { "%SystemRoot%", "{f38bf404-1d43-42f2-9305-67de0b28fc23}" }
    };

    public static void HashShortcut(string lnkPath)
    {
        SHCreateItemFromParsingName(lnkPath, null, typeof(IShellItem2).GUID, out var item);
        var item2 = (IShellItem2)item;

        PSGetPropertyKeyFromName("System.Link.TargetParsingPath", out var targetKey);
        string? targetPath;
        try { targetPath = item2.GetString(targetKey); } catch { targetPath = null; }

        PSGetPropertyKeyFromName("System.Link.Arguments", out var argsKey);
        string? arguments;
        try { arguments = item2.GetString(argsKey); } catch { arguments = null; }

        var blob = (GetGeneralizedPath(targetPath) + (arguments ?? string.Empty) +
                    "do not prehash links.  this should only be done by the user.").ToLowerInvariant();
        var inBytes = Encoding.Unicode.GetBytes(blob);
        var outBytes = new byte[inBytes.Length];
        HashData(inBytes, inBytes.Length, outBytes, outBytes.Length);
        var hash = BitConverter.ToUInt32(outBytes, 0);

        var guid = typeof(IPropertyStore).GUID;
        var store = item2.GetPropertyStore(GPS.READWRITE, ref guid);
        PSGetPropertyKeyFromName("System.Winx.Hash", out var hashKey);
        var pv = new PropVariant { VarType = VarEnum.VT_UI4, ulVal = hash };
        store.SetValue(ref hashKey, ref pv);
        store.Commit();

        Marshal.ReleaseComObject(store);
        Marshal.ReleaseComObject(item);
    }

    private static string GetGeneralizedPath(string? filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            return string.Empty;
        }
        foreach (var kv in GeneralizePathDic)
        {
            var dirPath = Environment.ExpandEnvironmentVariables(kv.Key);
            if (filePath.StartsWith(dirPath + "\\", StringComparison.OrdinalIgnoreCase))
            {
                return filePath.Replace(dirPath, kv.Value, StringComparison.OrdinalIgnoreCase);
            }
        }
        return filePath;
    }
}
