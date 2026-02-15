using System;
using System.Windows;
using Prism.Ioc;
using Prism.Navigation.Regions;

namespace BimTasksV2.Infrastructure
{
    /// <summary>
    /// Static helper for creating WPF windows with Prism region support.
    /// Equivalent to OldApp's PrismUtils.CreateWindow.
    /// </summary>
    public static class DialogHelper
    {
        /// <summary>
        /// Creates a Window of type TView, sets up Prism RegionManager, and returns it.
        /// </summary>
        public static TView CreateWindow<TView>() where TView : Window
        {
            TView window = Activator.CreateInstance<TView>();

            // Attach the Prism RegionManager so region navigation works inside the window
            var regionManager = ContainerLocator.Container.Resolve<IRegionManager>();
            RegionManager.SetRegionManager(window, regionManager);
            RegionManager.UpdateRegions();

            return window;
        }

        /// <summary>
        /// Creates a Window, sets Content to an instance of TView (UserControl), and returns it.
        /// </summary>
        public static Window CreateWindowWithContent<TView>(
            string title = "BimTasks",
            double width = 800,
            double height = 600,
            WindowStartupLocation startupLocation = WindowStartupLocation.CenterScreen)
            where TView : FrameworkElement, new()
        {
            var view = new TView();
            var window = new Window
            {
                Title = title,
                Width = width,
                Height = height,
                WindowStartupLocation = startupLocation,
                Content = view
            };

            var regionManager = ContainerLocator.Container.Resolve<IRegionManager>();
            RegionManager.SetRegionManager(window, regionManager);
            RegionManager.UpdateRegions();

            return window;
        }
    }
}
