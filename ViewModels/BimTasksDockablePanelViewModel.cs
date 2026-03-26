using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using BimTasksV2.Events;
using BimTasksV2.Services;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class TabInfo : BindableBase
    {
        public string Key { get; set; } = "";
        public string DisplayName { get; set; } = "";
        public bool CanClose { get; set; }
        public Action? CloseAction { get; set; }

        private bool _isActive;
        public bool IsActive
        {
            get => _isActive;
            set => SetProperty(ref _isActive, value);
        }
    }

    /// <summary>
    /// ViewModel for the BimTasks dockable panel.
    /// Subscribes to SwitchDockablePanelEvent and swaps the CurrentView content
    /// based on the view key (e.g. "FilterTree", "ElementCalculation").
    /// Also hosts the toolbar button groups inline.
    /// </summary>
    public class BimTasksDockablePanelViewModel : BindableBase
    {
        // Lazy-resolved: ICommandDispatcherService is registered in OnFirstIdle(),
        // but the panel is created during OnStartup() — before the service exists.
        private ICommandDispatcherService? _dispatcher;
        private ICommandDispatcherService Dispatcher =>
            _dispatcher ??= BimTasksV2.Infrastructure.ContainerLocator.Container
                .Resolve<ICommandDispatcherService>();

        private readonly Dictionary<string, FrameworkElement> _viewCache = new();
        private readonly Dictionary<string, string> _viewDisplayNames = new()
        {
            ["FilterTree"] = "Filter Tree",
            ["ElementCalculation"] = "Calculations",
            ["TrimCorners"] = "Trim Corners",
            ["ColorSwatch"] = "Color Swatch",
        };

        private string _panelTitle = "BimTasks Panel";
        public string PanelTitle
        {
            get => _panelTitle;
            set => SetProperty(ref _panelTitle, value);
        }

        private FrameworkElement? _currentView;
        public FrameworkElement? CurrentView
        {
            get => _currentView;
            set
            {
                SetProperty(ref _currentView, value);
                RaisePropertyChanged(nameof(EmptyStateVisibility));
            }
        }

        public System.Windows.Visibility EmptyStateVisibility =>
            _currentView == null ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        public System.Windows.Visibility TabStripVisibility =>
            Tabs.Count > 1 ? System.Windows.Visibility.Visible : System.Windows.Visibility.Collapsed;

        public ObservableCollection<TabInfo> Tabs { get; } = new();
        public DelegateCommand<string> SwitchTabCommand { get; }
        public DelegateCommand<string> CloseTabCommand { get; }

        public ObservableCollection<ToolbarGroupInfo> ToolbarGroups { get; }
        public DelegateCommand<string> ExecuteToolbarCommand { get; }

        public BimTasksDockablePanelViewModel()
        {
            ToolbarGroups = FloatingToolbarViewModel.BuildToolbarGroups();
            ExecuteToolbarCommand = new DelegateCommand<string>(OnExecuteToolbarCommand);
            SwitchTabCommand = new DelegateCommand<string>(OnSwitchTab);
            CloseTabCommand = new DelegateCommand<string>(OnCloseTab);

            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                .Subscribe(OnSwitchContent, ThreadOption.PublisherThread);
        }

        private void OnExecuteToolbarCommand(string? commandKey)
        {
            if (string.IsNullOrEmpty(commandKey)) return;

            Dispatcher.Enqueue(uiApp =>
            {
                try
                {
                    var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;

                    switch (commandKey)
                    {
                        case "FilterTree":
                            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                                .Publish("FilterTree");
                            return;

                        case "CalcAreaVolume":
                            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                                .Publish("ElementCalculation");
                            eventAgg.GetEvent<BimTasksEvents.CalculateElementsEvent>()
                                .Publish(null!);
                            return;
                    }

                    string handlerTypeName = $"BimTasksV2.Commands.Handlers.{commandKey}Handler";
                    var handlerType = typeof(BimTasksDockablePanelViewModel).Assembly.GetType(handlerTypeName);

                    if (handlerType != null)
                    {
                        var handler = Activator.CreateInstance(handlerType) as Commands.Infrastructure.ICommandHandler;
                        handler?.Execute(uiApp);
                        return;
                    }

                    Log.Warning("[Toolbar] Command not found: {Key}", commandKey);
                    Autodesk.Revit.UI.TaskDialog.Show("BimTasks",
                        $"Command '{commandKey}' is not yet available.");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[Toolbar] Failed to execute {Key}", commandKey);
                    Autodesk.Revit.UI.TaskDialog.Show("Error", $"Command failed:\n{ex.Message}");
                }
            });
        }

        private void OnSwitchTab(string? tabKey)
        {
            if (string.IsNullOrEmpty(tabKey)) return;
            if (!_viewCache.ContainsKey(tabKey)) return;
            ActivateTab(tabKey);
        }

        private void OnCloseTab(string? tabKey)
        {
            if (string.IsNullOrEmpty(tabKey)) return;

            var tab = Tabs.FirstOrDefault(t => t.Key == tabKey);
            if (tab == null) return;

            // Run cleanup callback if registered
            tab.CloseAction?.Invoke();

            _viewCache.Remove(tabKey);
            Tabs.Remove(tab);
            RaisePropertyChanged(nameof(TabStripVisibility));

            // Activate another tab or show empty state
            if (Tabs.Count > 0)
            {
                ActivateTab(Tabs.Last().Key);
            }
            else
            {
                PanelTitle = "BimTasks Panel";
                CurrentView = null;
            }
        }

        /// <summary>
        /// Registers a close action for a tab (e.g., clearing Revit overrides).
        /// Call this from the view/viewmodel that needs cleanup on tab close.
        /// </summary>
        public void RegisterTabCloseAction(string tabKey, Action closeAction)
        {
            var tab = Tabs.FirstOrDefault(t => t.Key == tabKey);
            if (tab != null)
            {
                tab.CanClose = true;
                tab.CloseAction = closeAction;
            }
        }

        private void ActivateTab(string tabKey)
        {
            foreach (var tab in Tabs)
                tab.IsActive = tab.Key == tabKey;

            var displayName = _viewDisplayNames.TryGetValue(tabKey, out var name) ? name : tabKey;
            PanelTitle = displayName;
            CurrentView = _viewCache[tabKey];
        }

        private void OnSwitchContent(string viewKey)
        {
            try
            {
                // If already cached, just activate
                if (_viewCache.ContainsKey(viewKey))
                {
                    ActivateTab(viewKey);
                    return;
                }

                // Create the view
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                FrameworkElement? view = null;

                switch (viewKey)
                {
                    case "FilterTree":
                        view = container.Resolve<Views.FilterTreeView>();
                        break;

                    case "ElementCalculation":
                        view = container.Resolve<Views.ElementCalculationView>();
                        break;

                    case "TrimCorners":
                        view = container.Resolve<Views.TrimCornersView>();
                        break;

                    case "ColorSwatch":
                        var colorSwatchView = new Views.ColorSwatchView();
                        view = colorSwatchView;
                        break;

                    default:
                        var viewType = Type.GetType($"BimTasksV2.Views.{viewKey}View");
                        if (viewType != null)
                        {
                            view = container.Resolve(viewType) as FrameworkElement;
                            if (!_viewDisplayNames.ContainsKey(viewKey))
                                _viewDisplayNames[viewKey] = viewKey;
                        }
                        else
                        {
                            Log.Warning("[DockablePanel] Unknown view key '{ViewKey}'", viewKey);
                            return;
                        }
                        break;
                }

                if (view == null) return;

                // Cache and create tab
                _viewCache[viewKey] = view;

                var displayName = _viewDisplayNames.TryGetValue(viewKey, out var dn) ? dn : viewKey;
                Tabs.Add(new TabInfo
                {
                    Key = viewKey,
                    DisplayName = displayName,
                    CanClose = false,
                });
                RaisePropertyChanged(nameof(TabStripVisibility));

                ActivateTab(viewKey);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[DockablePanel] Failed to switch to '{ViewKey}'", viewKey);
            }
        }
    }
}
