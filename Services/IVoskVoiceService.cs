using System;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Interface for the Vosk-based voice command recognition service.
    /// Manages microphone capture and dispatches recognized commands to Revit.
    /// </summary>
    public interface IVoskVoiceService : IDisposable
    {
        /// <summary>
        /// Whether the voice recognition is currently active.
        /// </summary>
        bool IsRunning { get; }

        /// <summary>
        /// Starts the voice recognition service (microphone capture + Vosk processing).
        /// </summary>
        void Start();

        /// <summary>
        /// Stops the voice recognition service and releases audio resources.
        /// </summary>
        void Stop();

        /// <summary>
        /// Toggles the voice recognition service on/off.
        /// </summary>
        void Toggle();
    }
}
