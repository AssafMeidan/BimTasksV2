using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using BimTasksV2.Events;
using BimTasksV2.ViewModels;
using Prism.Events;

namespace BimTasksV2.Views
{
    public partial class ColorSwatchView : UserControl
    {
        private readonly ColorSwatchViewModel _vm;

        public ColorSwatchView()
        {
            _vm = new ColorSwatchViewModel();
            DataContext = _vm;
            InitializeComponent();

            // Subscribe to initialization event from handler
            var eventAgg = BimTasksV2.Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.InitializeColorSwatchEvent>()
                .Subscribe(OnInitialize, ThreadOption.PublisherThread);

            // Register tab close cleanup
            Loaded += (s, e) =>
            {
                // Walk up to find the dockable panel ViewModel and register close action
                var panelVm = FindPanelViewModel();
                panelVm?.RegisterTabCloseAction("ColorSwatch", () => _vm.OnTabClosed());
            };
        }

        private void OnInitialize(Autodesk.Revit.UI.UIApplication uiApp)
        {
            _vm.Initialize(uiApp);
        }

        private void SwatchList_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (sender is not ListView lv || lv.SelectedItem is not SwatchRowItem row) return;
            _vm.SelectElementsCommand.Execute(row);
        }

        private BimTasksDockablePanelViewModel? FindPanelViewModel()
        {
            DependencyObject? current = this;
            while (current != null)
            {
                if (current is FrameworkElement fe && fe.DataContext is BimTasksDockablePanelViewModel vm)
                    return vm;
                current = LogicalTreeHelper.GetParent(current)
                          ?? VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        private void Swatch_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is not FrameworkElement fe || fe.DataContext is not SwatchRowItem row) return;
            if (DataContext is not ColorSwatchViewModel vm) return;
            e.Handled = true; // prevent ListViewItem from capturing the click

            // Create preset color popup
            var popup = new Popup
            {
                PlacementTarget = fe,
                Placement = PlacementMode.Right,
                StaysOpen = false,
                AllowsTransparency = true,
            };

            var panel = new WrapPanel
            {
                Width = 165,
                Background = Brushes.White,
                Margin = new Thickness(0),
            };

            var border = new Border
            {
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0xCC, 0xCC, 0xCC)),
                BorderThickness = new Thickness(1),
                Padding = new Thickness(6),
                Background = Brushes.White,
                Child = new StackPanel
                {
                    Children =
                    {
                        panel,
                    }
                }
            };

            // Add preset color swatches
            foreach (var presetColor in ColorSwatchViewModel.PresetColors)
            {
                var swatch = new Border
                {
                    Width = 24,
                    Height = 24,
                    Margin = new Thickness(2),
                    CornerRadius = new CornerRadius(2),
                    BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(0x99, 0x99, 0x99)),
                    BorderThickness = new Thickness(1),
                    Background = new SolidColorBrush(presetColor),
                    Cursor = Cursors.Hand,
                };
                var color = presetColor; // capture for closure
                swatch.MouseLeftButtonDown += (s, args) =>
                {
                    vm.UpdateSwatchColor(row, color);
                    popup.IsOpen = false;
                };
                panel.Children.Add(swatch);
            }

            // Add "Custom..." button
            var customBtn = new Button
            {
                Content = "Custom...",
                Margin = new Thickness(2, 6, 2, 0),
                Padding = new Thickness(4, 3, 4, 3),
                FontSize = 11,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                Cursor = Cursors.Hand,
            };
            customBtn.Click += (s, args) =>
            {
                popup.IsOpen = false;

                var dlg = new System.Windows.Forms.ColorDialog
                {
                    Color = System.Drawing.Color.FromArgb(row.Color.R, row.Color.G, row.Color.B),
                    FullOpen = true,
                };
                if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    var c = dlg.Color;
                    vm.UpdateSwatchColor(row,
                        System.Windows.Media.Color.FromRgb(c.R, c.G, c.B));
                }
            };
            ((StackPanel)border.Child).Children.Add(customBtn);

            popup.Child = border;
            popup.IsOpen = true;
        }
    }
}
