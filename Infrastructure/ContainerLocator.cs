using Prism.Events;
using Prism.Ioc;

namespace BimTasksV2.Infrastructure
{
    /// <summary>
    /// Static wrapper around the Prism DI container.
    /// Provides global access to the container and event aggregator
    /// from code that cannot use constructor injection (e.g., proxy commands).
    /// </summary>
    public static class ContainerLocator
    {
        public static IContainerExtension Container { get; private set; } = null!;
        public static IEventAggregator EventAggregator { get; private set; } = null!;

        /// <summary>
        /// Sets the global container and resolves the event aggregator.
        /// Called once from BimTasksBootstrapper.Initialize().
        /// </summary>
        public static void SetContainer(IContainerExtension container)
        {
            Container = container;
            EventAggregator = container.Resolve<IEventAggregator>();
        }
    }
}
