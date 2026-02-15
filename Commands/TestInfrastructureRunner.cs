using System;
using System.Text;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Events;
using BimTasksV2.Services;
using BimTasksV2.ViewModels;
using Prism.Events;
using Prism.Ioc;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Comprehensive diagnostic runner that validates the entire BimTasksV2 infrastructure.
    /// Tests DI container, Revit access, event flow, view models, and dockable panel.
    /// Uses fully-qualified BimTasksV2.Infrastructure.ContainerLocator because the
    /// Commands namespace has its own Infrastructure sub-namespace.
    /// </summary>
    public class TestInfrastructureRunner : ICommandHandler
    {
        private int _passed;
        private int _failed;
        private readonly StringBuilder _report = new();

        public void Execute(UIApplication uiApp)
        {
            _passed = 0;
            _failed = 0;
            _report.Clear();

            _report.AppendLine("=== BimTasksV2 Infrastructure Diagnostics ===");
            _report.AppendLine();

            TestDIContainerAndServices();
            TestRevitDocumentAccess(uiApp);
            TestFilterTreePipeline();
            TestDockablePanelAndViewSwitching();
            TestEventFlow();

            _report.AppendLine();
            _report.AppendLine("=== Summary ===");
            _report.AppendLine($"Passed: {_passed}  |  Failed: {_failed}  |  Total: {_passed + _failed}");

            string title = _failed == 0
                ? $"All {_passed} Tests Passed"
                : $"{_failed} of {_passed + _failed} Tests Failed";

            TaskDialog.Show($"BimTasksV2 Diagnostics - {title}", _report.ToString());
        }

        // =====================================================================
        // Section 1: DI Container & Services
        // =====================================================================

        private void TestDIContainerAndServices()
        {
            _report.AppendLine("--- Section 1: DI Container & Services ---");

            // Test ContainerLocator itself
            TestStep("ContainerLocator.Container is set",
                () => BimTasksV2.Infrastructure.ContainerLocator.Container != null);

            TestStep("ContainerLocator.EventAggregator is set",
                () => BimTasksV2.Infrastructure.ContainerLocator.EventAggregator != null);

            // Test IEventAggregator resolution
            TestStep("Resolve IEventAggregator",
                () => BimTasksV2.Infrastructure.ContainerLocator.Container.Resolve<IEventAggregator>() != null);

            // Test IRevitContextService
            TestStep("Resolve IRevitContextService",
                () => BimTasksV2.Infrastructure.ContainerLocator.Container.Resolve<IRevitContextService>() != null);

            // Test ICommandDispatcherService
            TestStep("Resolve ICommandDispatcherService",
                () => BimTasksV2.Infrastructure.ContainerLocator.Container.Resolve<ICommandDispatcherService>() != null);

            _report.AppendLine();
        }

        // =====================================================================
        // Section 2: Revit Document Access
        // =====================================================================

        private void TestRevitDocumentAccess(UIApplication uiApp)
        {
            _report.AppendLine("--- Section 2: Revit Document Access ---");

            TestStep("UIApplication is valid",
                () => uiApp != null);

            TestStep("ActiveUIDocument is available",
                () => uiApp?.ActiveUIDocument != null);

            if (uiApp?.ActiveUIDocument != null)
            {
                var doc = uiApp.ActiveUIDocument.Document;
                TestStep("Document is not null",
                    () => doc != null);

                TestStep("Document title is accessible",
                    () =>
                    {
                        string title = doc?.Title;
                        _report.AppendLine($"       Document: {title ?? "(untitled)"}");
                        return !string.IsNullOrEmpty(title);
                    });

                TestStep("Document path is accessible",
                    () =>
                    {
                        string path = doc?.PathName;
                        _report.AppendLine($"       Path: {(string.IsNullOrEmpty(path) ? "(not saved)" : path)}");
                        return true; // path can be empty for unsaved docs
                    });
            }
            else
            {
                _report.AppendLine("  [SKIP] No active document — skipping document tests.");
            }

            // Test that IRevitContextService has been populated
            try
            {
                var ctx = BimTasksV2.Infrastructure.ContainerLocator.Container.Resolve<IRevitContextService>();
                TestStep("IRevitContextService.UIApplication is set",
                    () => ctx?.UIApplication != null);
            }
            catch
            {
                RecordFail("IRevitContextService.UIApplication is set", "Could not resolve service");
            }

            _report.AppendLine();
        }

        // =====================================================================
        // Section 3: Filter Tree Pipeline
        // =====================================================================

        private void TestFilterTreePipeline()
        {
            _report.AppendLine("--- Section 3: Filter Tree Pipeline ---");

            try
            {
                var ea = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                TestStep("Can get ResetFilterTreeEvent",
                    () => ea.GetEvent<BimTasksEvents.ResetFilterTreeEvent>() != null);

                TestStep("Can publish ResetFilterTreeEvent",
                    () =>
                    {
                        ea.GetEvent<BimTasksEvents.ResetFilterTreeEvent>().Publish(null);
                        return true;
                    });
            }
            catch (Exception ex)
            {
                RecordFail("Filter Tree Pipeline", ex.Message);
            }

            _report.AppendLine();
        }

        // =====================================================================
        // Section 4: Dockable Panel & View Switching
        // =====================================================================

        private void TestDockablePanelAndViewSwitching()
        {
            _report.AppendLine("--- Section 4: Dockable Panel & View Switching ---");

            TestStep("BimTasksApp.DockablePaneId is set",
                () => BimTasksApp.DockablePaneId != null);

            try
            {
                var ea = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                TestStep("Can get SwitchDockablePanelEvent",
                    () => ea.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>() != null);

                TestStep("Can publish SwitchDockablePanelEvent (FilterTreeView)",
                    () =>
                    {
                        ea.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>().Publish("FilterTreeView");
                        return true;
                    });
            }
            catch (Exception ex)
            {
                RecordFail("Dockable Panel", ex.Message);
            }

            _report.AppendLine();
        }

        // =====================================================================
        // Section 5: Event Flow
        // =====================================================================

        private void TestEventFlow()
        {
            _report.AppendLine("--- Section 5: Event Flow ---");

            try
            {
                var ea = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
                bool received = false;

                // Subscribe with PublisherThread to ensure synchronous delivery
                var token = ea.GetEvent<BimTasksEvents.CalculateElementsEvent>()
                    .Subscribe(
                        _ => received = true,
                        ThreadOption.PublisherThread);

                // Publish
                ea.GetEvent<BimTasksEvents.CalculateElementsEvent>().Publish(null);

                TestStep("CalculateElementsEvent delivered via PublisherThread",
                    () => received);

                // Unsubscribe
                ea.GetEvent<BimTasksEvents.CalculateElementsEvent>().Unsubscribe(token);

                // Test ToggleFloatingToolbarEvent
                bool toolbarReceived = false;
                var toolbarToken = ea.GetEvent<BimTasksEvents.ToggleFloatingToolbarEvent>()
                    .Subscribe(
                        _ => toolbarReceived = true,
                        ThreadOption.PublisherThread);

                ea.GetEvent<BimTasksEvents.ToggleFloatingToolbarEvent>().Publish(null);

                TestStep("ToggleFloatingToolbarEvent delivered via PublisherThread",
                    () => toolbarReceived);

                ea.GetEvent<BimTasksEvents.ToggleFloatingToolbarEvent>().Unsubscribe(toolbarToken);
            }
            catch (Exception ex)
            {
                RecordFail("Event Flow", ex.Message);
            }

            _report.AppendLine();
        }

        // =====================================================================
        // Helpers
        // =====================================================================

        private void TestStep(string description, Func<bool> test)
        {
            try
            {
                if (test())
                {
                    _report.AppendLine($"  [PASS] {description}");
                    _passed++;
                }
                else
                {
                    _report.AppendLine($"  [FAIL] {description}");
                    _failed++;
                }
            }
            catch (Exception ex)
            {
                _report.AppendLine($"  [FAIL] {description} -- {ex.Message}");
                _failed++;
            }
        }

        private void RecordFail(string description, string reason)
        {
            _report.AppendLine($"  [FAIL] {description} -- {reason}");
            _failed++;
        }
    }
}
