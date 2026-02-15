using System.Windows;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Pure WPF dialog for changing wall height (top/bottom offsets).
    /// Exposes properties that the handler reads after ShowDialog().
    /// </summary>
    public partial class ChangeWallHeightDialog : Window
    {
        /// <summary>Whether to modify the top height.</summary>
        public bool IsAddToTop { get; set; }

        /// <summary>Whether to modify the bottom height.</summary>
        public bool IsAddToBottom { get; set; } = true;

        /// <summary>Amount to add to top height, in centimeters.</summary>
        public double AddToTop { get; set; } = 25;

        /// <summary>Amount to add to bottom offset, in centimeters.</summary>
        public double AddToBottom { get; set; } = 30;

        /// <summary>True when user clicked OK (not Cancel).</summary>
        public bool Accepted { get; private set; }

        public ChangeWallHeightDialog()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Accepted = true;
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Accepted = false;
            DialogResult = false;
            Close();
        }
    }
}
