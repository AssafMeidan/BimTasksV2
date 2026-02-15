using System;
using System.Windows;
using BimTasksV2.Events;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    /// <summary>
    /// ViewModel for the BimTasks dockable panel.
    /// Subscribes to SwitchDockablePanelEvent and swaps the CurrentView content
    /// based on the view key (e.g. "FilterTree", "ElementCalculation").
    /// </summary>
    public class BimTasksDockablePanelViewModel : BindableBase
    {
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

        public Visibility EmptyStateVisibility =>
            _currentView == null ? Visibility.Visible : Visibility.Collapsed;

        public BimTasksDockablePanelViewModel()
        {
            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                .Subscribe(OnSwitchContent, ThreadOption.PublisherThread);
        }

        private void OnSwitchContent(string viewKey)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;

                switch (viewKey)
                {
                    case "FilterTree":
                        PanelTitle = "Filter Tree";
                        CurrentView = container.Resolve<Views.FilterTreeView>();
                        break;

                    case "ElementCalculation":
                        PanelTitle = "Element Calculations";
                        CurrentView = container.Resolve<Views.ElementCalculationView>();
                        break;

                    default:
                        // Try to resolve by full type name within the Views namespace
                        var viewType = Type.GetType($"BimTasksV2.Views.{viewKey}View");
                        if (viewType != null)
                        {
                            PanelTitle = viewKey;
                            CurrentView = container.Resolve(viewType) as FrameworkElement;
                        }
                        else
                        {
                            Log.Warning("[DockablePanel] Unknown view key '{ViewKey}'", viewKey);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[DockablePanel] Failed to switch to '{ViewKey}'", viewKey);
            }
        }
    }
}
