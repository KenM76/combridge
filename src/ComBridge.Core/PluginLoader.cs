using System.Reflection;
using System.Runtime.Loader;

namespace ComBridge.Core;

/// <summary>
/// Discovers plugins under <c>{exeDir}/plugins/{pluginName}/*.dll</c>.
/// Each plugin DLL must export at least one type implementing <see cref="IComBridgePlugin"/>.
/// </summary>
public static class PluginLoader
{
    /// <summary>
    /// Default plugin root directory: <c>plugins/</c> next to the host executable.
    /// Resolved via <see cref="AppContext.BaseDirectory"/>, so it works whether
    /// the host is run from its publish folder, a staging folder, or a relative
    /// path.
    /// </summary>
    public static string DefaultPluginRoot =>
        Path.Combine(AppContext.BaseDirectory, "plugins");

    /// <summary>
    /// Discover and instantiate all plugins under <paramref name="pluginRoot"/>
    /// (defaults to <see cref="DefaultPluginRoot"/>). For each subdirectory:
    /// loads <c>ComBridge.Plugins.&lt;dirName&gt;.dll</c> (preferred) or any
    /// <c>*Plugin*.dll</c> fallback into a per-folder
    /// <see cref="AssemblyLoadContext"/>, then enumerates exported types that
    /// implement <see cref="IComBridgePlugin"/> and instantiates them via the
    /// parameterless constructor.
    /// </summary>
    /// <param name="pluginRoot">
    /// Root directory to scan. <c>null</c> = <see cref="DefaultPluginRoot"/>.
    /// </param>
    /// <returns>
    /// The list of successfully-instantiated plugins, in directory enumeration
    /// order. Plugins whose DLLs fail to load (missing dependencies, type-load
    /// errors, constructor exceptions) are skipped with a diagnostic written to
    /// <see cref="Console.Error"/>, NOT thrown — the host stays usable even if
    /// one plugin is broken.
    /// </returns>
    /// <remarks>
    /// Per-folder <see cref="AssemblyLoadContext"/> instances allow plugins to
    /// ship transitive dependencies that don't collide with each other or with
    /// the host. The inner <c>PluginLoadContext.Load</c> prefers already-loaded
    /// assemblies from the Default context (so Core, BCL, and Roslyn aren't
    /// duplicated per plugin), then falls back to
    /// <see cref="AssemblyDependencyResolver"/>, then to side-by-side DLLs in
    /// the plugin folder.
    /// </remarks>
    public static IReadOnlyList<IComBridgePlugin> LoadAll(string? pluginRoot = null)
    {
        var root = pluginRoot ?? DefaultPluginRoot;
        var found = new List<IComBridgePlugin>();
        if (!Directory.Exists(root)) return found;

        foreach (var dir in Directory.EnumerateDirectories(root))
        {
            // Each plugin lives in its own folder so we can load deps with an
            // AssemblyLoadContext rooted at that folder.
            var name = Path.GetFileName(dir);
            var primary = Path.Combine(dir, $"ComBridge.Plugins.{name}.dll");
            if (!File.Exists(primary))
            {
                // Fall back to any *.Plugin*.dll in the folder.
                primary = Directory.EnumerateFiles(dir, "*.dll")
                    .FirstOrDefault(f => Path.GetFileName(f).Contains("Plugin", StringComparison.OrdinalIgnoreCase))
                    ?? "";
                if (primary == "") continue;
            }

            try
            {
                var ctx = new PluginLoadContext(dir);
                var asm = ctx.LoadFromAssemblyPath(primary);
                foreach (var t in asm.GetTypes())
                {
                    if (!typeof(IComBridgePlugin).IsAssignableFrom(t)) continue;
                    if (t.IsAbstract || t.IsInterface) continue;
                    if (Activator.CreateInstance(t) is IComBridgePlugin p) found.Add(p);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"[plugin-load] {primary}: {ex.Message}");
            }
        }
        return found;
    }

    private sealed class PluginLoadContext : AssemblyLoadContext
    {
        private readonly AssemblyDependencyResolver _resolver;
        private readonly string _dir;

        public PluginLoadContext(string dir) : base(isCollectible: false)
        {
            _dir = dir;
            // AssemblyDependencyResolver wants a path to the primary assembly,
            // but a folder works for our fallback lookup below.
            var probe = Directory.EnumerateFiles(dir, "*.dll").FirstOrDefault() ?? dir;
            _resolver = new AssemblyDependencyResolver(probe);
        }

        protected override Assembly? Load(AssemblyName name)
        {
            // Prefer assemblies already loaded into the default context (Core, BCL).
            var existing = Default.Assemblies.FirstOrDefault(a => a.GetName().Name == name.Name);
            if (existing is not null) return existing;

            var resolved = _resolver.ResolveAssemblyToPath(name);
            if (resolved is not null) return LoadFromAssemblyPath(resolved);

            var sideBySide = Path.Combine(_dir, name.Name + ".dll");
            if (File.Exists(sideBySide)) return LoadFromAssemblyPath(sideBySide);
            return null;
        }
    }
}
