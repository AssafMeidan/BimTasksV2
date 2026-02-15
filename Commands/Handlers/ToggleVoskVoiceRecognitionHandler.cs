using System;
using Autodesk.Revit.UI;
using BimTasksV2.Commands.Infrastructure;
using BimTasksV2.Services;
using Prism.Ioc;
using Serilog;

namespace BimTasksV2.Commands.Handlers
{
    /// <summary>
    /// Toggles the Vosk voice recognition service on/off.
    /// Resolves IVoskVoiceService from DI and calls Toggle().
    /// Uses fully-qualified BimTasksV2.Infrastructure.ContainerLocator because the
    /// Commands namespace has its own Infrastructure sub-namespace.
    /// </summary>
    public class ToggleVoskVoiceRecognitionHandler : ICommandHandler
    {
        public void Execute(UIApplication uiApp)
        {
            try
            {
                var container = BimTasksV2.Infrastructure.ContainerLocator.Container;
                var voiceService = container.Resolve<IVoskVoiceService>();

                if (voiceService == null)
                {
                    TaskDialog.Show("Voice Recognition",
                        "Voice recognition service is not available.\n\n" +
                        "Make sure Vosk and NAudio packages are installed and the Vosk model is present.");
                    return;
                }

                bool wasRunning = voiceService.IsRunning;
                voiceService.Toggle();

                string status = voiceService.IsRunning
                    ? "Voice recognition started.\n\nSpeak commands like: wall, door, copy, undo, section..."
                    : "Voice recognition stopped.";

                TaskDialog.Show("Voice Recognition", status);

                Log.Information("Voice recognition toggled: {PreviousState} -> {CurrentState}",
                    wasRunning ? "Running" : "Stopped",
                    voiceService.IsRunning ? "Running" : "Stopped");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to toggle voice recognition");
                TaskDialog.Show("Voice Recognition Error",
                    $"Failed to toggle voice recognition:\n\n{ex.Message}");
            }
        }
    }
}
