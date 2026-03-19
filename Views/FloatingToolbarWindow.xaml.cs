using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Input;
using System.Windows.Shell;

namespace BimTasksV2.Views
{
    /// <summary>
    /// Singleton floating toolbar window with custom chrome, resize via WindowChrome,
    /// opacity transitions, and position/size persistence.
    /// Closing is intercepted — the window hides instead.
    /// </summary>
    public partial class FloatingToolbarWindow : Window
    {
        public FloatingToolbarWindow()
        {
            InitializeComponent();

            // WindowChrome provides resize borders on the borderless transparent window
            WindowChrome.SetWindowChrome(this, new WindowChrome
            {
                CaptionHeight = 0,
                ResizeBorderThickness = new Thickness(8),
                GlassFrameThickness = new Thickness(0),
                CornerRadius = new CornerRadius(0),
                UseAeroCaptionButtons = false
            });

            DataContext = new ViewModels.FloatingToolbarViewModel();

            // Save state whenever the window becomes hidden
            IsVisibleChanged += (s, e) =>
            {
                if (e.NewValue is false)
                    SaveCurrentState();
            };

            // Listen for orientation changes to swap dimensions
            if (DataContext is ViewModels.FloatingToolbarViewModel vm)
                vm.OrientationChanged += OnOrientationChanged;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is not ViewModels.FloatingToolbarViewModel vm)
                return;

            var (left, top, width, height) = vm.LoadSettings();

            // Restore size
            if (width.HasValue && width.Value > MinWidth)
                Width = width.Value;
            if (height.HasValue && height.Value > MinHeight)
                Height = height.Value;

            // Restore position (validate on-screen)
            if (left.HasValue && top.HasValue)
            {
                double screenW = SystemParameters.VirtualScreenWidth;
                double screenH = SystemParameters.VirtualScreenHeight;
                double screenL = SystemParameters.VirtualScreenLeft;
                double screenT = SystemParameters.VirtualScreenTop;

                if (left.Value >= screenL && left.Value < screenL + screenW - 50
                    && top.Value >= screenT && top.Value < screenT + screenH - 50)
                {
                    Left = left.Value;
                    Top = top.Value;
                }
            }
        }

        private void Header_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 1)
                DragMove();
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            e.Cancel = true;
            Hide();
        }

        private void OnOrientationChanged()
        {
            // Swap width and height for a natural orientation transition
            (Width, Height) = (Height, Width);
            InvalidateMeasure();
            UpdateLayout();
        }

        // Opacity transitions: full opacity when active/hovered, reduced when unfocused
        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            Opacity = 1.0;
        }

        protected override void OnDeactivated(EventArgs e)
        {
            base.OnDeactivated(e);
            if (!IsMouseOver)
                Opacity = 0.85;
        }

        protected override void OnMouseEnter(MouseEventArgs e)
        {
            base.OnMouseEnter(e);
            Opacity = 1.0;
        }

        protected override void OnMouseLeave(MouseEventArgs e)
        {
            base.OnMouseLeave(e);
            if (!IsActive)
                Opacity = 0.85;
        }

        private void SaveCurrentState()
        {
            (DataContext as ViewModels.FloatingToolbarViewModel)?.SaveSettings(Left, Top, Width, Height);
        }
    }
}
