using System.Diagnostics;
using Serilog;

namespace BimTasksV2.Logging
{
    /// <summary>
    /// Captures WPF data binding errors and routes them to Serilog.
    /// WPF binding failures are normally silent; this listener makes them visible
    /// in the log so they can be diagnosed during development.
    /// </summary>
    public class SerilogTraceListener : TraceListener
    {
        public override void Write(string? message)
        {
            // Intentionally empty — binding errors come through WriteLine
        }

        public override void WriteLine(string? message)
        {
            if (message != null && message.Contains("BindingExpression"))
            {
                Log.Warning("WPF BINDING ERROR: {Message}", message);
            }
        }

        /// <summary>
        /// Registers this listener on PresentationTraceSources.DataBindingSource
        /// so all WPF binding errors are captured.
        /// </summary>
        public static void Register()
        {
            PresentationTraceSources.DataBindingSource.Listeners.Add(new SerilogTraceListener());
            PresentationTraceSources.DataBindingSource.Switch.Level = SourceLevels.Warning;
        }
    }
}
