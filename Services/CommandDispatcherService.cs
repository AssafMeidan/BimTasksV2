using Autodesk.Revit.UI;
using Serilog;
using System;
using System.Collections.Concurrent;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Bridges WPF UI context (toolbar buttons, dockable panel) to Revit API context
    /// via ExternalEvent. Callers enqueue Action&lt;UIApplication&gt; lambdas from any thread;
    /// Revit executes them on its main thread during the next idle cycle.
    /// Must be created on Revit's main thread (ExternalEvent.Create requirement).
    /// </summary>
    public sealed class CommandDispatcherService : IExternalEventHandler, ICommandDispatcherService
    {
        private readonly ConcurrentQueue<Action<UIApplication>> _actionQueue = new();
        private readonly ExternalEvent _externalEvent;
        private readonly object _lock = new();
        private readonly UIApplication _uiApplication;

        public CommandDispatcherService(UIApplication uiApplication)
        {
            _uiApplication = uiApplication ?? throw new ArgumentNullException(nameof(uiApplication));
            _externalEvent = ExternalEvent.Create(this);
        }

        /// <summary>
        /// Queue an action for execution in Revit API context.
        /// Safe to call from any thread (WPF button click, voice command, etc.).
        /// </summary>
        public void Enqueue(Action<UIApplication> action)
        {
            if (action == null) return;

            lock (_lock)
            {
                _actionQueue.Enqueue(action);
                _externalEvent.Raise();
            }
        }

        /// <summary>
        /// Called by Revit on the main thread. Drains and executes all queued actions.
        /// </summary>
        public void Execute(UIApplication app)
        {
            while (_actionQueue.TryDequeue(out var action))
            {
                try
                {
                    action(app);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "CommandDispatcher: action failed");
                }
            }
        }

        public string GetName() => "BimTasksCommandDispatcher";
    }
}
