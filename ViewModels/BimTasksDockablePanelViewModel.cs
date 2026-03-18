using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using BimTasksV2.Events;
using Prism.Commands;
using Prism.Events;
using Prism.Ioc;
using Prism.Mvvm;
using Serilog;

namespace BimTasksV2.ViewModels
{
    public class DockableTabItem : BindableBase
    {
        public string Key { get; set; } = "";
        public string Title { get; set; } = "";
        public FrameworkElement? Content { get; set; }
    }

    /// <summary>
    /// ViewModel for the BimTasks dockable panel.
    /// Manages a collection of tabs, each hosting a different view.
    /// Re-opening an existing view activates its tab instead of creating a duplicate.
    /// </summary>
    public class BimTasksDockablePanelViewModel : BindableBase
    {
        public ObservableCollection<DockableTabItem> Tabs { get; } = new();

        private DockableTabItem? _selectedTab;
        public DockableTabItem? SelectedTab
        {
            get => _selectedTab;
            set
            {
                SetProperty(ref _selectedTab, value);
                RaisePropertyChanged(nameof(EmptyStateVisibility));
            }
        }

        public Visibility EmptyStateVisibility =>
            Tabs.Count == 0 ? Visibility.Visible : Visibility.Collapsed;

        public ICommand CloseTabCommand { get; }

        public BimTasksDockablePanelViewModel()
        {
            CloseTabCommand = new DelegateCommand<DockableTabItem>(OnCloseTab);

            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.SwitchDockablePanelEvent>()
                .Subscribe(OnSwitchContent, ThreadOption.PublisherThread);
        }

        private void OnCloseTab(DockableTabItem? tab)
        {
            if (tab == null) return;

            var index = Tabs.IndexOf(tab);
            Tabs.Remove(tab);

            if (Tabs.Count > 0)
            {
                // Select the tab at the same position, or the last one
                SelectedTab = Tabs[Math.Min(index, Tabs.Count - 1)];
            }
            else
            {
                SelectedTab = null;
            }

            RaisePropertyChanged(nameof(EmptyStateVisibility));
        }

        private void OnSwitchContent(string viewKey)
        {
            try
            {
                // If tab already exists for this view key, just activate it
                var existing = Tabs.FirstOrDefault(t => t.Key == viewKey);
                if (existing != null)
                {
                    SelectedTab = existing;
                    return;
                }

                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                string title;
                FrameworkElement? view;

                switch (viewKey)
                {
                    case "FilterTree":
                        title = "Filter Tree";
                        view = container.Resolve<Views.FilterTreeView>();
                        break;

                    case "ElementCalculation":
                        title = "Element Calculations";
                        view = container.Resolve<Views.ElementCalculationView>();
                        break;

                    case "FixSplitCorners":
                        title = "Fix Split Corners";
                        view = container.Resolve<Views.FixSplitCornersView>();
                        break;

                    case "ColorCodeByParameter":
                        title = "Color Code by Parameter";
                        view = container.Resolve<Views.ColorCodeByParameterView>();
                        break;

                    default:
                        var viewType = Type.GetType($"BimTasksV2.Views.{viewKey}View");
                        if (viewType != null)
                        {
                            title = viewKey;
                            view = container.Resolve(viewType) as FrameworkElement;
                        }
                        else
                        {
                            Log.Warning("[DockablePanel] Unknown view key '{ViewKey}'", viewKey);
                            return;
                        }
                        break;
                }

                if (view == null) return;

                var tab = new DockableTabItem
                {
                    Key = viewKey,
                    Title = title,
                    Content = view
                };

                Tabs.Add(tab);
                SelectedTab = tab;
                RaisePropertyChanged(nameof(EmptyStateVisibility));
            }
            catch (Exception ex)
            {
                Log.Error(ex, "[DockablePanel] Failed to switch to '{ViewKey}'", viewKey);
            }
        }
    }
}
