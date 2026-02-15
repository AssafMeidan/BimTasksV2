using System.Windows;
using System.Windows.Controls;
using BimTasksV2.ViewModels;

namespace BimTasksV2.Views
{
    public partial class FilterTreeView : UserControl
    {
        public FilterTreeView()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Explicit handler to ensure CheckBox state propagates to the TreeNodeViewModel.
        /// WPF binding inside HierarchicalDataTemplate may not update the source reliably
        /// in the isolated AssemblyLoadContext.
        /// </summary>
        private void CheckBox_Toggle(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb && cb.DataContext is TreeNodeViewModel node)
                node.IsChecked = cb.IsChecked ?? false;
        }
    }
}
