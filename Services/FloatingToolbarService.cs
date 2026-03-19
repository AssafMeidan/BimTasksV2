using System;
using BimTasksV2.Events;
using BimTasksV2.Views;
using Prism.Events;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Manages the singleton floating toolbar window.
    /// Subscribes to ToggleFloatingToolbarEvent to show/hide the toolbar.
    /// </summary>
    public class FloatingToolbarService
    {
        private FloatingToolbarWindow? _window;

        public FloatingToolbarService()
        {
            var eventAgg = Infrastructure.ContainerLocator.EventAggregator;
            eventAgg.GetEvent<BimTasksEvents.ToggleFloatingToolbarEvent>()
                .Subscribe(_ => Toggle(), ThreadOption.PublisherThread);

            Log.Information("FloatingToolbarService: Initialized and listening for toggle events.");
        }

        public void Toggle()
        {
            try
            {
                if (_window == null)
                    _window = new FloatingToolbarWindow();

                if (_window.IsVisible)
                    _window.Hide();
                else
                    _window.Show();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "FloatingToolbarService: Toggle failed");
            }
        }
    }
}
