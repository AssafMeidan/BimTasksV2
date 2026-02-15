using System.ComponentModel;
using System.Windows;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Singleton topmost floating toolbar window.
    /// Closing is intercepted; the window hides instead of closing (singleton pattern).
    /// </summary>
    public partial class FloatingToolbarWindow : Window
    {
        public FloatingToolbarWindow()
        {
            InitializeComponent();
            DataContext = new ViewModels.FloatingToolbarViewModel();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            // Don't close, just hide (singleton pattern)
            e.Cancel = true;
            Hide();
        }
    }
}
