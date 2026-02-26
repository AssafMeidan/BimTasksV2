using System;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using OfficeOpenXml;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Prism.Navigation.Regions;
using Prism.Navigation.Regions.Behaviors;
using Prism.Container.Unity;
using Serilog;
using BimTasksV2.Logging;
using BimTasksV2.Services;

namespace BimTasksV2.Infrastructure
{
    /// <summary>
    /// Prism 9 / Unity manual bootstrapper for BimTasksV2.
    /// Runs inside the isolated BimTasksLoadContext to avoid DLL conflicts
    /// with Revit's default ALC (Enscape, Dynamo, etc.).
    ///
    /// Lifecycle:
    ///   1. Initialize()        — called from BimTasksApp.OnStartup via reflection
    ///   2. RegisterDockablePane — called during OnStartup (must happen before first idle)
    ///   3. OnFirstIdle()       — called once from Revit Idling event
    ///   4. Shutdown()          — called from BimTasksApp.OnShutdown via reflection
    /// </summary>
    public class BimTasksBootstrapper
    {
        private IContainerExtension _container = null!;
        private VoskVoiceCommandService? _voiceService;

        /// <summary>
        /// DockablePaneId GUID for the BimTasks dockable panel.
        /// </summary>
        private static readonly Guid DockablePaneGuid = new("B2C3D4E5-F6A7-8901-BCDE-FA2345678902");

        #region Phase 1: Initialize (called from OnStartup)

        /// <summary>
        /// Creates the Unity container, registers all Prism infrastructure,
        /// application services, and views for navigation.
        /// </summary>
        public void Initialize()
        {
            // Initialize logging first so everything after this is captured
            SerilogConfig.Initialize();
            SerilogTraceListener.Register();
            Log.Information("BimTasksBootstrapper.Initialize() starting...");

            // Set EPPlus license (must happen before any EPPlus usage)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Create Prism 9 Unity container
            _container = new UnityContainerExtension();

            // Register the container itself (Prism convention)
            _container.RegisterInstance(_container);

            // Register Prism infrastructure singletons
            RegisterPrismInfrastructure(_container);

            // Register application services
            RegisterServices(_container);

            // Register views for region navigation
            RegisterViews(_container);

            // Set the global static container locator EARLY — Prism internals
            // (RegionAdapterMappings, behaviors, etc.) need ContainerLocator.Container
            // to be set before they can resolve dependencies.
            ContainerLocator.SetContainer(_container);

            // Wire up ViewModelLocator (view -> VM auto-resolve)
            ViewModelLocationProvider.SetDefaultViewModelFactory((view, type) =>
            {
                return _container.Resolve(type);
            });

            // Configure region adapter mappings
            var regionAdapterMappings = _container.Resolve<RegionAdapterMappings>();
            if (regionAdapterMappings != null)
            {
                regionAdapterMappings.RegisterMapping<Selector, SelectorRegionAdapter>();
                regionAdapterMappings.RegisterMapping<ItemsControl, ItemsControlRegionAdapter>();
                regionAdapterMappings.RegisterMapping<ContentControl, ContentControlRegionAdapter>();
            }

            // Configure region behaviors
            var regionBehaviorFactory = _container.Resolve<IRegionBehaviorFactory>();
            if (regionBehaviorFactory != null)
            {
                regionBehaviorFactory.AddIfMissing<BindRegionContextToDependencyObjectBehavior>(
                    BindRegionContextToDependencyObjectBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<RegionActiveAwareBehavior>(
                    RegionActiveAwareBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<SyncRegionContextWithHostBehavior>(
                    SyncRegionContextWithHostBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<RegionManagerRegistrationBehavior>(
                    RegionManagerRegistrationBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<RegionMemberLifetimeBehavior>(
                    RegionMemberLifetimeBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<ClearChildViewsRegionBehavior>(
                    ClearChildViewsRegionBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<AutoPopulateRegionBehavior>(
                    AutoPopulateRegionBehavior.BehaviorKey);
                regionBehaviorFactory.AddIfMissing<DestructibleRegionBehavior>(
                    DestructibleRegionBehavior.BehaviorKey);
            }

            Log.Information("BimTasksBootstrapper.Initialize() completed successfully.");
        }

        #endregion Phase 1: Initialize (called from OnStartup)

        #region Phase 2: Register Dockable Pane (called during OnStartup)

        /// <summary>
        /// Creates the DockablePaneId and registers the BimTasksDockablePanel with Revit.
        /// Must be called during IExternalApplication.OnStartup — Revit does not allow
        /// pane registration after startup completes.
        /// </summary>
        public DockablePaneId RegisterDockablePane(UIControlledApplication app)
        {
            var paneId = new DockablePaneId(DockablePaneGuid);

            try
            {
                var panel = new Views.BimTasksDockablePanel();
                app.RegisterDockablePane(paneId, "BimTasks", panel);
                Log.Information("Dockable pane registered: {PaneId}", DockablePaneGuid);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to register dockable pane");
            }

            return paneId;
        }

        #endregion Phase 2: Register Dockable Pane (called during OnStartup)

        #region Phase 3: OnFirstIdle (called from Revit Idling event)

        /// <summary>
        /// Performs initialization that requires the Revit main thread:
        /// - Sets UIApplication on RevitContextService
        /// - Creates CommandDispatcherService (requires ExternalEvent.Create on main thread)
        /// - Creates VoskVoiceCommandService singleton
        /// </summary>
        public void OnFirstIdle(UIApplication uiApp)
        {
            Log.Information("BimTasksBootstrapper.OnFirstIdle() starting...");

            // Set the UIApplication on the context service
            var contextService = _container.Resolve<IRevitContextService>();
            contextService.UIApplication = uiApp;

            // Create CommandDispatcherService on Revit's main thread
            // (ExternalEvent.Create must be called from the Revit thread)
            var dispatcher = new CommandDispatcherService(uiApp);
            _container.RegisterInstance<ICommandDispatcherService>(dispatcher);
            _container.RegisterInstance<CommandDispatcherService>(dispatcher);

            // Create VoskVoiceCommandService singleton
            try
            {
                _voiceService = new VoskVoiceCommandService(uiApp);
                _container.RegisterInstance<IVoskVoiceService>(_voiceService);
                Log.Information("VoskVoiceCommandService created and registered.");
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Failed to create VoskVoiceCommandService (voice commands unavailable)");
            }

            Log.Information("BimTasksBootstrapper.OnFirstIdle() completed.");
        }

        #endregion Phase 3: OnFirstIdle (called from Revit Idling event)

        #region Phase 4: Shutdown

        /// <summary>
        /// Disposes resources. Called from BimTasksApp.OnShutdown via reflection.
        /// </summary>
        public void Shutdown()
        {
            Log.Information("BimTasksBootstrapper.Shutdown() starting...");

            try
            {
                _voiceService?.Dispose();
                _voiceService = null;
            }
            catch (Exception ex)
            {
                Log.Warning(ex, "Error disposing VoskVoiceCommandService");
            }

            Log.Information("BimTasksBootstrapper.Shutdown() completed.");
            Log.CloseAndFlush();
        }

        #endregion Phase 4: Shutdown

        #region Registration Helpers

        /// <summary>
        /// Registers Prism 9 infrastructure singletons required for regions,
        /// navigation, dialogs, and events.
        /// NOTE: Prism 9 removed ILoggerFacade, IServiceLocator, and FinalizeExtension().
        /// </summary>
        private static void RegisterPrismInfrastructure(IContainerRegistry registry)
        {
            registry.RegisterSingleton<IRegionManager, RegionManager>();
            registry.RegisterSingleton<IEventAggregator, EventAggregator>();
            registry.RegisterSingleton<RegionAdapterMappings>();
            registry.RegisterSingleton<IRegionViewRegistry, RegionViewRegistry>();
            registry.RegisterSingleton<IRegionBehaviorFactory, RegionBehaviorFactory>();
            registry.RegisterSingleton<IRegionNavigationContentLoader, RegionNavigationContentLoader>();
            registry.Register<IRegionNavigationJournalEntry, RegionNavigationJournalEntry>();
            registry.Register<IRegionNavigationJournal, RegionNavigationJournal>();
            registry.Register<IRegionNavigationService, RegionNavigationService>();
        }

        /// <summary>
        /// Registers all BimTasksV2 application services.
        /// Services that require UIApplication (CommandDispatcherService, VoskVoiceCommandService)
        /// are registered later in OnFirstIdle().
        /// </summary>
        private static void RegisterServices(IContainerRegistry registry)
        {
            // Core services (available immediately)
            registry.RegisterSingleton<IRevitContextService, RevitContextService>();

            // Calculation / data services
            registry.RegisterSingleton<IElementCalculationService, ElementCalculationService>();
            registry.RegisterSingleton<IUniformatDataService, UniformatDataService>();
            registry.RegisterSingleton<IRevitUniformatWriter, RevitUniformatWriter>();
            registry.RegisterSingleton<IDataExtractorService, DataExtractorService>();
            registry.RegisterSingleton<IValidationScheduleService, ValidationScheduleService>();

            // Excel roundtrip service (singleton — holds cached schedule mappings)
            registry.RegisterSingleton<ScheduleExcelRoundtripService>();
        }

        /// <summary>
        /// Registers views for Prism region navigation.
        /// Each view is registered with its type name as the navigation key.
        /// </summary>
        private static void RegisterViews(IContainerRegistry registry)
        {
            // Dockable panel child views
            registry.RegisterForNavigation<Views.FilterTreeView>();
            registry.RegisterForNavigation<Views.ElementCalculationView>();
            registry.RegisterForNavigation<Views.UniformatWindowView>();
            registry.RegisterForNavigation<Views.CopyCategoryFromLinkView>();
        }

        #endregion Registration Helpers
    }
}