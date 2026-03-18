using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using BimTasksV2.ViewModels;

namespace BimTasksV2.Views
{
    public partial class ColorCodeByParameterView : UserControl
    {
        public ColorCodeByParameterView() { InitializeComponent(); }

        private void SwatchButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is LegendItem item
                && DataContext is ColorCodeByParameterViewModel vm)
            {
                vm.EditingLegendItem = item;
                ColorPickerPopup.PlacementTarget = btn;
                ColorPickerPopup.IsOpen = true;
            }
        }

        private void PaletteColor_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.DataContext is Color color
                && DataContext is ColorCodeByParameterViewModel vm)
            {
                vm.ApplyColorToItem(color);
                ColorPickerPopup.IsOpen = false;
            }
        }
    }
}
