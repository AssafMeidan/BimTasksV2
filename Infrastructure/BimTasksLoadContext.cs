using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.Loader;

namespace BimTasksV2.Infrastructure
{
    /// <summary>
    /// Custom AssemblyLoadContext that isolates BimTasks dependencies (Prism, Unity, Vosk)
    /// from the default Revit context to prevent DLL conflicts with Enscape and Dynamo.
    /// </summary>
    public class BimTasksLoadContext : AssemblyLoadContext
    {
        private readonly AssemblyDependencyResolver _resolver;
        private readonly string _pluginDirectory;
        private readonly List<string> _loadLog = new();

        /// <summary>
        /// Assembly names that should always resolve from the default (Revit/WPF) context.
        /// These are framework assemblies that Revit hosts and must be shared.
        /// </summary>
        private static readonly HashSet<string> _defaultContextAssemblies = new(StringComparer.OrdinalIgnoreCase)
        {
            // Revit API
            "RevitAPI",
            "RevitAPIUI",
            "RevitAPIIFC",
            "RevitAPIMacros",
            // WPF / UI framework — must share with Revit's hosting
            "PresentationFramework",
            "PresentationCore",
            "WindowsBase",
            "System.Xaml",
            "UIAutomationProvider",
            "UIAutomationTypes",
            // .NET runtime — always from default
            "System.Runtime",
            "System.Private.CoreLib",
            "netstandard",
            "mscorlib",
        };

        /// <summary>
        /// Prefixes for assemblies that should always resolve from the default context.
        /// </summary>
        private static readonly string[] _defaultContextPrefixes = new[]
        {
            "System.",
            "Microsoft.Win32.",
            "Microsoft.CSharp",
            "Microsoft.VisualBasic",
        };

        public IReadOnlyList<string> LoadLog => _loadLog;

        public BimTasksLoadContext(string pluginPath) : base(name: "BimTasksIsolated", isCollectible: false)
        {
            _pluginDirectory = Path.GetDirectoryName(pluginPath)
                ?? throw new ArgumentException("Invalid plugin path", nameof(pluginPath));
            _resolver = new AssemblyDependencyResolver(pluginPath);
        }

        protected override Assembly? Load(AssemblyName assemblyName)
        {
            string name = assemblyName.Name ?? string.Empty;

            // Check explicit default-context assemblies
            if (_defaultContextAssemblies.Contains(name))
            {
                Log($"DEFAULT  {name} (explicit allow-list)");
                return null; // null = fall back to default context
            }

            // Check prefix-based default-context assemblies
            foreach (var prefix in _defaultContextPrefixes)
            {
                if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    Log($"DEFAULT  {name} (prefix: {prefix})");
                    return null;
                }
            }

            // Try resolving from our isolated plugin directory
            string? resolvedPath = _resolver.ResolveAssemblyToPath(assemblyName);
            if (resolvedPath != null)
            {
                Log($"ISOLATED {name} -> {resolvedPath}");
                return LoadFromAssemblyPath(resolvedPath);
            }

            // Also check directly in the plugin directory (fallback for deps not in .deps.json)
            string directPath = Path.Combine(_pluginDirectory, $"{name}.dll");
            if (File.Exists(directPath))
            {
                Log($"ISOLATED {name} -> {directPath} (direct)");
                return LoadFromAssemblyPath(directPath);
            }

            // Not found in our context — fall back to default
            Log($"DEFAULT  {name} (not found in isolated context)");
            return null;
        }

        protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
        {
            // Try resolving through the dependency resolver first
            string? resolvedPath = _resolver.ResolveUnmanagedDllToPath(unmanagedDllName);
            if (resolvedPath != null)
            {
                Log($"NATIVE   {unmanagedDllName} -> {resolvedPath}");
                return LoadUnmanagedDllFromPath(resolvedPath);
            }

            // Try common native library naming patterns in our plugin directory
            string[] candidates = new[]
            {
                Path.Combine(_pluginDirectory, unmanagedDllName),
                Path.Combine(_pluginDirectory, $"{unmanagedDllName}.dll"),
                Path.Combine(_pluginDirectory, "runtimes", "win-x64", "native", $"{unmanagedDllName}.dll"),
                Path.Combine(_pluginDirectory, "runtimes", "win-x64", "native", unmanagedDllName),
            };

            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    Log($"NATIVE   {unmanagedDllName} -> {candidate}");
                    return LoadUnmanagedDllFromPath(candidate);
                }
            }

            Log($"NATIVE   {unmanagedDllName} -> NOT FOUND (falling back to default)");
            return IntPtr.Zero;
        }

        private void Log(string message)
        {
            string entry = $"[{DateTime.Now:HH:mm:ss.fff}] {message}";
            _loadLog.Add(entry);
        }

        /// <summary>
        /// Returns a formatted report of all assembly load decisions for diagnostics.
        /// </summary>
        public string GetLoadReport()
        {
            return string.Join(Environment.NewLine, _loadLog);
        }
    }
}
