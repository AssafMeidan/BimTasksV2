using Prism.Events;

namespace BimTasksV2.Events
{
    public class FixSplitCornersPayload
    {
        public int WallsSplit { get; set; }
        public int TotalReplacements { get; set; }
    }

    /// <summary>
    /// Prism PubSubEvent definitions for BimTasksV2.
    /// Events are published/subscribed via the IEventAggregator resolved from DI.
    /// </summary>
    public static class BimTasksEvents
    {
        /// <summary>
        /// Raised to reset the FilterTree view to its initial state.
        /// </summary>
        public class ResetFilterTreeEvent : PubSubEvent<object> { }

        /// <summary>
        /// Raised to trigger element calculation in the ElementCalculationView.
        /// </summary>
        public class CalculateElementsEvent : PubSubEvent<object> { }

        /// <summary>
        /// Raised to switch the dockable panel content.
        /// Payload is the view type name (e.g., "FilterTreeView", "ElementCalculationView").
        /// </summary>
        public class SwitchDockablePanelEvent : PubSubEvent<string> { }

        /// <summary>
        /// Raised after Split Wall completes to signal the dockable panel
        /// that corner fix data is ready.
        /// </summary>
        public class FixSplitCornersReadyEvent : PubSubEvent<FixSplitCornersPayload> { }

        /// <summary>
        /// Raised to toggle the floating toolbar window visibility.
        /// </summary>
        public class ToggleFloatingToolbarEvent : PubSubEvent<object> { }
    }
}
