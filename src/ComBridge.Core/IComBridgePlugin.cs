using System.Runtime.InteropServices;
using Microsoft.CodeAnalysis;

namespace ComBridge.Core;

/// <summary>
/// Contract every per-app plugin implements. The host (combridge.exe) discovers
/// plugins from the plugins/ folder and dispatches CLI args to the matching plugin.
/// </summary>
public interface IComBridgePlugin
{
    /// <summary>Name used on the CLI, e.g. "solidworks", "excel".</summary>
    string Name { get; }

    /// <summary>Friendly description for `combridge list-plugins`.</summary>
    string Description { get; }

    /// <summary>
    /// ProgIDs to attempt when attaching to a running instance via the ROT,
    /// in priority order. Example: ["SldWorks.Application"].
    /// </summary>
    string[] ProgIds { get; }

    /// <summary>
    /// If no running instance is found and <see cref="AllowCreateNew"/> is true,
    /// the host calls <c>Activator.CreateInstance(Type.GetTypeFromProgID(...))</c>
    /// using the first ProgID.
    /// </summary>
    bool AllowCreateNew { get; }

    /// <summary>
    /// Type of the globals object exposed to user scripts (e.g. SwGlobals).
    /// Roslyn binds script identifiers like <c>swApp</c> against this type.
    /// </summary>
    Type GlobalsType { get; }

    /// <summary>
    /// Build the globals instance from a connected COM root object.
    /// </summary>
    object CreateGlobals(object comRoot);

    /// <summary>
    /// Assemblies to add as MetadataReferences when compiling user scripts.
    /// Typically the plugin's own assembly + any interop DLLs it uses.
    /// </summary>
    IEnumerable<MetadataReference> ScriptReferences { get; }

    /// <summary>Namespaces auto-imported in user scripts.</summary>
    IEnumerable<string> ScriptUsings { get; }

    /// <summary>
    /// Plugin-specific commands (in addition to the built-in <c>run-script</c>).
    /// Examples: SW <c>checkpoint</c>, Excel <c>dump-sheet</c>.
    /// </summary>
    IEnumerable<IBridgeCommand> Commands { get; }

    /// <summary>
    /// Optional: extract a PID and title from a connected COM root, for the
    /// <c>list-sessions</c> picker and <c>--session</c> selector. Default
    /// returns nothing so existing plugins still compile.
    /// </summary>
    (int? Pid, string? Title) DescribeInstance(object comRoot) => (null, null);

    /// <summary>
    /// Regex patterns matched against each <c>IMoniker.GetDisplayName</c>
    /// in the ROT. A plugin instance is discovered when its app registers a
    /// moniker whose display name matches ANY of these patterns.
    /// <para>
    /// Why this exists separately from <see cref="ProgIds"/>: real COM servers
    /// rarely register their ProgID literally as the ROT display name. SolidWorks
    /// uses <c>SolidWorks_PID_&lt;pid&gt;</c>; Excel uses per-document file
    /// monikers (<c>.xlsx</c>, <c>.xlsm</c>, …). The plugin author knows the
    /// actual format; the host doesn't.
    /// </para>
    /// <para>
    /// Default: matches the bare ProgID(s) as substrings (legacy behavior — works
    /// for COM servers that register their ProgID directly). Override to provide
    /// proper patterns when your app uses a different ROT moniker format.
    /// </para>
    /// Patterns are full .NET regex, evaluated case-insensitively.
    /// </summary>
    IEnumerable<string> RotMonikerPatterns
        => ProgIds.Select(p => System.Text.RegularExpressions.Regex.Escape(p));

    /// <summary>
    /// Convert a raw ROT-bound object into the "root" type
    /// <see cref="CreateGlobals"/> expects. Default = identity.
    /// <para>
    /// Use this when your patterns match per-document monikers (e.g. Excel
    /// matches <c>.xlsx</c> file monikers and ascends <c>Workbook.Application</c>
    /// to get the parent app). Return null to skip this binding (e.g. the
    /// moniker bound to an unexpected type).
    /// </para>
    /// <para>
    /// Apps reachable via <c>oleaut32!GetActiveObject</c> already return the
    /// Application root directly, so this hook only fires on ROT-walk results.
    /// </para>
    /// </summary>
    object? TryExtractRoot(object monikerBound) => monikerBound;

    /// <summary>
    /// OS platforms this plugin supports. <see cref="PluginLoader"/> silently
    /// skips plugins whose <see cref="SupportedPlatforms"/> doesn't include the
    /// current OS (checked via <see cref="OperatingSystem"/>).
    /// <para>
    /// Default = Windows only — matches the behavior of every plugin written
    /// against combridge v0.2.x, which all use Windows COM via
    /// <see cref="RotHelper"/> and <c>oleaut32!GetActiveObject</c> and have no
    /// non-Windows code path. New cross-platform plugins (e.g. AppleScript-based
    /// Office plugins for macOS, UNO-based LibreOffice on Linux) override this
    /// to declare their actual support: <c>new[] { OSPlatform.OSX }</c>,
    /// <c>new[] { OSPlatform.OSX, OSPlatform.Linux }</c>, etc.
    /// </para>
    /// </summary>
    IReadOnlyCollection<OSPlatform> SupportedPlatforms => new[] { OSPlatform.Windows };

    /// <summary>
    /// Discover all live instances of this plugin's target app on the current
    /// OS, MRU-sorted (most-recently-focused first where applicable). Returns
    /// the same shape as <see cref="SessionPicker.Enumerate"/> — a list of
    /// <c>(Root, Info)</c> tuples ready for <c>list-sessions</c> display or
    /// <c>--session</c> selector resolution.
    /// <para>
    /// <b>Default implementation</b> (Windows only, available when this Core
    /// assembly is built for <c>net10.0-windows</c>): delegates to
    /// <see cref="SessionPicker.Enumerate"/>, which combines ROT-moniker pattern
    /// walking (using <see cref="RotMonikerPatterns"/>) with
    /// <c>oleaut32!GetActiveObject</c> fallback (using <see cref="ProgIds"/>),
    /// applies <see cref="TryExtractRoot"/>, dedupes by PID, drops dead bindings,
    /// and sorts by desktop Z-order.
    /// </para>
    /// <para>
    /// <b>Non-Windows plugins MUST override this</b> with platform-native
    /// discovery (e.g. <c>osascript</c> + <c>tell application "System Events"</c>
    /// on macOS, UNO bridge enumeration on Linux). The default throws
    /// <see cref="PlatformNotSupportedException"/> on any non-Windows build,
    /// making the "you forgot to override" case loud at runtime.
    /// </para>
    /// </summary>
    List<(object Root, SessionInfo Info)> FindSessions()
    {
#if WINDOWS
        return SessionPicker.Enumerate(this);
#else
        throw new PlatformNotSupportedException(
            $"Plugin '{Name}' didn't override FindSessions(), and the default " +
            $"Windows-COM implementation isn't available on this OS. " +
            $"Non-Windows plugins must implement their own discovery.");
#endif
    }
}
