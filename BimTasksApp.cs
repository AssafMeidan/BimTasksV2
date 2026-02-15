using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;

namespace BimTasksV2
{
    /// <summary>
    /// Revit entry point. Thin proxy in default ALC — must NOT reference Prism/Unity/Serilog.
    /// Creates the isolated BimTasksLoadContext and delegates to BimTasksBootstrapper via reflection.
    /// </summary>
    public class BimTasksApp : IExternalApplication
    {
        internal static Infrastructure.BimTasksLoadContext? LoadContext { get; private set; }

        /// <summary>
        /// DockablePaneId for the BimTasks panel, set during startup via reflection from bootstrapper.
        /// </summary>
        public static DockablePaneId? DockablePaneId { get; internal set; }

        private static object? _bootstrapper;
        private static Type? _bootstrapperType;
        private static UIControlledApplication? _controlledApp;

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                _controlledApp = application;
                string assemblyPath = Assembly.GetExecutingAssembly().Location;

                // Create isolated load context
                LoadContext = new Infrastructure.BimTasksLoadContext(assemblyPath);

                // Load assembly into isolated context
                var isolatedAssembly = LoadContext.LoadFromAssemblyPath(assemblyPath);

                // Bootstrap Prism 9 / Unity in isolated context via reflection
                _bootstrapperType = isolatedAssembly.GetType("BimTasksV2.Infrastructure.BimTasksBootstrapper")
                    ?? throw new InvalidOperationException("BimTasksBootstrapper type not found in isolated assembly.");

                _bootstrapper = Activator.CreateInstance(_bootstrapperType);

                var initMethod = _bootstrapperType.GetMethod("Initialize", BindingFlags.Public | BindingFlags.Instance)
                    ?? throw new InvalidOperationException("Initialize method not found on BimTasksBootstrapper.");
                initMethod.Invoke(_bootstrapper, null);

                // Create ribbon tab and all panels
                Ribbon.RibbonBuilder.Build(application, assemblyPath);

                // Register dockable pane (must happen during OnStartup)
                var registerPaneMethod = _bootstrapperType.GetMethod("RegisterDockablePane", BindingFlags.Public | BindingFlags.Instance)
                    ?? throw new InvalidOperationException("RegisterDockablePane method not found on BimTasksBootstrapper.");
                var paneIdResult = registerPaneMethod.Invoke(_bootstrapper, new object[] { application });
                if (paneIdResult is DockablePaneId dpId)
                    DockablePaneId = dpId;

                // Subscribe to first idle for Revit-thread-only initialization
                application.Idling += OnFirstIdle;

                return Result.Succeeded;
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                TaskDialog.Show("BimTasksV2 Error",
                    $"Startup failed:\n{ex.InnerException.Message}\n\n{ex.InnerException.StackTrace}");
                return Result.Failed;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("BimTasksV2 Error",
                    $"Startup failed:\n{ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            try
            {
                if (_bootstrapper != null && _bootstrapperType != null)
                {
                    var shutdownMethod = _bootstrapperType.GetMethod("Shutdown", BindingFlags.Public | BindingFlags.Instance);
                    shutdownMethod?.Invoke(_bootstrapper, null);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"BimTasksV2 shutdown error: {ex.Message}");
            }

            return Result.Succeeded;
        }

        private static void OnFirstIdle(object? sender, IdlingEventArgs e)
        {
            // One-shot: unsubscribe immediately
            _controlledApp!.Idling -= OnFirstIdle;

            try
            {
                if (sender is not UIApplication uiApp)
                    return;

                var method = _bootstrapperType!.GetMethod("OnFirstIdle", BindingFlags.Public | BindingFlags.Instance)
                    ?? throw new InvalidOperationException("OnFirstIdle method not found on BimTasksBootstrapper.");
                method.Invoke(_bootstrapper, new object[] { uiApp });
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                TaskDialog.Show("BimTasksV2 Error",
                    $"First-idle init failed:\n{ex.InnerException.Message}");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("BimTasksV2 Error",
                    $"First-idle init failed:\n{ex.Message}");
            }
        }

    }
}
