using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComBridge.Core;

/// <summary>
/// COM interop layer for hosting Microsoft Script engines (VBScript,
/// JScript, etc.) in-process via the documented <c>IActiveScript</c>
/// interface family. All declarations here are the canonical
/// activscp.h shapes — only what <see cref="VbScriptEngine"/> actually
/// calls is declared.
/// </summary>
/// <remarks>
/// <para>
/// References:
/// </para>
/// <list type="bullet">
///   <item>Microsoft "Active Script Hosting Architecture" docs (legacy
///         MSDN, archived). The interfaces are stable since IE4 / 1998
///         and have not changed shape since.</item>
///   <item>The 64-bit pointer variants (<c>IActiveScriptParse64</c>,
///         <c>IActiveScriptParseProcedure64</c>) are required on 64-bit
///         hosts. The non-suffixed names (<c>IActiveScriptParse</c>) are
///         32-bit only. .NET 10 on x64 must use the 64 variants.</item>
///   <item>The script engines ship as part of Windows (vbscript.dll,
///         jscript.dll). They are NOT redistributable — relying on the
///         in-box install is the supported path. VBScript was formally
///         deprecated in 2024; the host will stop working when Microsoft
///         removes the engine from a future Windows release.</item>
/// </list>
/// </remarks>
internal static class ActiveScriptInterop
{
    // ============================================================
    // Constants
    // ============================================================

    /// <summary>
    /// CLSID for the in-box VBScript engine, used with
    /// <see cref="Type.GetTypeFromCLSID(System.Guid)"/> +
    /// <see cref="Activator.CreateInstance(System.Type)"/> to instantiate.
    /// </summary>
    public static readonly Guid CLSID_VBScript = new("B54F3741-5B07-11cf-A4B0-00AA004A55E8");

    /// <summary>
    /// CLSID for the in-box JScript engine. Not used in v0.8.0; declared
    /// here so a future JScript host can be added in <30 lines by
    /// reusing every other type in this file.
    /// </summary>
    public static readonly Guid CLSID_JScript = new("F414C260-6AC0-11CF-B6D1-00AA00BBBB58");

    // SCRIPTSTATE values (see activscp.h enum tagSCRIPTSTATE).
    public const uint SCRIPTSTATE_UNINITIALIZED = 0;
    public const uint SCRIPTSTATE_INITIALIZED   = 5;
    public const uint SCRIPTSTATE_STARTED       = 1;
    public const uint SCRIPTSTATE_CONNECTED     = 2;
    public const uint SCRIPTSTATE_DISCONNECTED  = 3;
    public const uint SCRIPTSTATE_CLOSED        = 4;

    // AddNamedItem flags. SCRIPTITEM_GLOBALMEMBERS makes the item's
    // members (properties/methods) accessible WITHOUT the item name as a
    // prefix — so VBScript can call WScript.Echo() (with prefix) and
    // also call methods of an unnamed-members object directly. We use
    // ISVISIBLE without GLOBALMEMBERS for the named globals (swApp,
    // swDoc, etc.) — the user writes the name explicitly.
    public const uint SCRIPTITEM_ISVISIBLE      = 0x00000002;
    public const uint SCRIPTITEM_ISSOURCE       = 0x00000004;
    public const uint SCRIPTITEM_GLOBALMEMBERS  = 0x00000008;
    public const uint SCRIPTITEM_NOCODE         = 0x00000400;
    public const uint SCRIPTITEM_CODEONLY       = 0x00000200;

    // GetItemInfo mask bits. We honor IUNKNOWN (return the IDispatch
    // pointer) and ignore ITYPEINFO (return null) — late-binding is
    // sufficient for SW automation, and supplying ITypeInfo would add
    // significant marshalling complexity for no real gain.
    public const uint SCRIPTINFO_IUNKNOWN   = 0x00000001;
    public const uint SCRIPTINFO_ITYPEINFO  = 0x00000002;

    // ParseScriptText flags. ISVISIBLE makes top-level VBScript
    // statements run immediately at parse time (the "global script" body),
    // which is the standard interpretation of a .vbs file. Without it
    // we'd need to call into a named entry-point Sub explicitly.
    public const uint SCRIPTTEXT_ISVISIBLE          = 0x00000002;
    public const uint SCRIPTTEXT_ISPERSISTENT       = 0x00000040;
    public const uint SCRIPTTEXT_HOSTMANAGESSOURCE  = 0x00000080;

    // ============================================================
    // Structs
    // ============================================================

    /// <summary>
    /// Mirrors <c>EXCEPINFO</c> from oaidl.h. Returned by
    /// <see cref="IActiveScriptParse64.ParseScriptText"/> on a syntax
    /// error and by <see cref="IActiveScriptError.GetExceptionInfo"/> on
    /// a runtime error.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct EXCEPINFO
    {
        public ushort wCode;
        public ushort wReserved;
        [MarshalAs(UnmanagedType.BStr)] public string? bstrSource;
        [MarshalAs(UnmanagedType.BStr)] public string? bstrDescription;
        [MarshalAs(UnmanagedType.BStr)] public string? bstrHelpFile;
        public uint dwHelpContext;
        public IntPtr pvReserved;
        public IntPtr pfnDeferredFillIn;
        public int scode;
    }

    // ============================================================
    // Interfaces
    // ============================================================

    /// <summary>
    /// <c>IActiveScript</c> — the engine. Created via
    /// <c>CoCreateInstance(CLSID_VBScript)</c>. The host sets a script
    /// site (the callback object), adds named items (globals), then
    /// transitions the engine state to CONNECTED to run the script.
    /// </summary>
    [ComImport, Guid("BB1A2AE1-A4F9-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScript
    {
        [PreserveSig] int SetScriptSite([In, MarshalAs(UnmanagedType.Interface)] IActiveScriptSite pSite);
        [PreserveSig] int GetScriptSite(ref Guid riid, out IntPtr ppvObject);
        [PreserveSig] int SetScriptState(uint state);
        [PreserveSig] int GetScriptState(out uint state);
        [PreserveSig] int Close();
        [PreserveSig] int AddNamedItem(
            [MarshalAs(UnmanagedType.LPWStr)] string pstrName,
            uint dwFlags);
        [PreserveSig] int AddTypeLib(ref Guid rguidTypeLib, uint dwMajor, uint dwMinor, uint dwFlags);
        [PreserveSig] int GetScriptDispatch(
            [MarshalAs(UnmanagedType.LPWStr)] string? pstrItemName,
            out IntPtr ppdisp);
        [PreserveSig] int GetCurrentScriptThreadID(out uint pstidThread);
        [PreserveSig] int GetScriptThreadID(uint dwWin32ThreadId, out uint pstidThread);
        [PreserveSig] int GetScriptThreadState(uint stidThread, out uint pstateThread);
        [PreserveSig] int InterruptScriptThread(uint stidThread, ref EXCEPINFO excepInfo, uint dwFlags);
        [PreserveSig] int Clone(out IntPtr ppscript);
    }

    /// <summary>
    /// <c>IActiveScriptParse64</c> — the 64-bit parse interface. The
    /// non-suffixed <c>IActiveScriptParse</c> uses 32-bit DWORD context
    /// cookies which break on x64. Both VBScript and JScript implement
    /// the 64 variant on 64-bit Windows.
    /// </summary>
    [ComImport, Guid("C7EF7658-E1EE-480E-97EA-D52CB4D76D17")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptParse64
    {
        [PreserveSig] int InitNew();
        [PreserveSig] int AddScriptlet(
            [MarshalAs(UnmanagedType.LPWStr)] string pstrDefaultName,
            [MarshalAs(UnmanagedType.LPWStr)] string pstrCode,
            [MarshalAs(UnmanagedType.LPWStr)] string pstrItemName,
            [MarshalAs(UnmanagedType.LPWStr)] string pstrSubItemName,
            [MarshalAs(UnmanagedType.LPWStr)] string pstrEventName,
            [MarshalAs(UnmanagedType.LPWStr)] string pstrDelimiter,
            ulong dwSourceContextCookie,
            uint ulStartingLineNumber,
            uint dwFlags,
            [MarshalAs(UnmanagedType.BStr)] out string pbstrName,
            out EXCEPINFO pexcepinfo);
        [PreserveSig] int ParseScriptText(
            [MarshalAs(UnmanagedType.LPWStr)] string pstrCode,
            [MarshalAs(UnmanagedType.LPWStr)] string? pstrItemName,
            [MarshalAs(UnmanagedType.IUnknown)] object? punkContext,
            [MarshalAs(UnmanagedType.LPWStr)] string? pstrDelimiter,
            ulong dwSourceContextCookie,
            uint ulStartingLineNumber,
            uint dwFlags,
            IntPtr pvarResult,
            out EXCEPINFO pexcepinfo);
    }

    /// <summary>
    /// <c>IActiveScriptSite</c> — the callback object the engine uses to
    /// resolve named items, report errors, and notify state changes. The
    /// host implements this and passes it via
    /// <see cref="IActiveScript.SetScriptSite"/>.
    /// </summary>
    [ComImport, Guid("DB01A1E3-A42B-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptSite
    {
        [PreserveSig] int GetLCID(out uint plcid);
        [PreserveSig] int GetItemInfo(
            [MarshalAs(UnmanagedType.LPWStr)] string pstrName,
            uint dwReturnMask,
            [MarshalAs(UnmanagedType.IUnknown)] out object? ppiunkItem,
            IntPtr ppti);
        [PreserveSig] int GetDocVersionString([MarshalAs(UnmanagedType.BStr)] out string? pbstrVersion);
        [PreserveSig] int OnScriptTerminate([In] ref object pvarResult, [In] ref EXCEPINFO pexcepinfo);
        [PreserveSig] int OnStateChange(uint state);
        [PreserveSig] int OnScriptError([In, MarshalAs(UnmanagedType.Interface)] IActiveScriptError pscripterror);
        [PreserveSig] int OnEnterScript();
        [PreserveSig] int OnLeaveScript();
    }

    /// <summary>
    /// <c>IActiveScriptError</c> — handed to the host's
    /// <see cref="IActiveScriptSite.OnScriptError"/> callback. Lets us
    /// extract the EXCEPINFO, source position, and the offending line of
    /// text for human-readable error reporting.
    /// </summary>
    [ComImport, Guid("EAE1BA61-A4ED-11cf-8F20-00805F2CD064")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IActiveScriptError
    {
        [PreserveSig] int GetExceptionInfo(out EXCEPINFO pexcepinfo);
        [PreserveSig] int GetSourcePosition(
            out uint pdwSourceContext,
            out uint pulLineNumber,
            out int plCharacterPosition);
        [PreserveSig] int GetSourceLineText([MarshalAs(UnmanagedType.BStr)] out string? pbstrSourceLine);
    }
}
