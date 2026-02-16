using System;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text.Json;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using NAudio.Wave;
using Serilog;

namespace BimTasksV2.Services
{
    /// <summary>
    /// Provides voice command recognition using Vosk offline speech recognition.
    /// Captures audio via NAudio, processes with Vosk, and dispatches recognized
    /// commands to Revit via the UIApplication.Idling event (thread-safe).
    /// Uses System.Text.Json (not Newtonsoft.Json).
    /// </summary>
    public class VoskVoiceCommandService : IVoskVoiceService
    {
        #region Fields

        private WaveInEvent _waveIn;
        private Vosk.VoskRecognizer _recognizer;
        private Vosk.Model _model;
        private readonly UIApplication _uiApp;
        private readonly ConcurrentQueue<string> _commandQueue = new();
        private readonly object _lock = new();

        private bool _isRunning;
        private bool _isHandlerAttached;
        private bool _disposed;

        // Duplicate command prevention
        private string _lastCommandText;
        private DateTime _lastCommandTime;
        private const int DuplicateThresholdMs = 800;

        // Audio configuration
        private const int SampleRate = 16000;
        private const int Channels = 1;

        #endregion

        #region Properties

        public bool IsRunning => _isRunning;

        #endregion

        #region Events

        /// <summary>
        /// Fired when a voice command is recognized (before execution).
        /// </summary>
        public event Action<string> OnCommandRecognized;

        /// <summary>
        /// Fired when a command is successfully executed.
        /// </summary>
        public event Action<string> OnCommandExecuted;

        #endregion

        #region Constructor

        public VoskVoiceCommandService(UIApplication uiApp)
        {
            _uiApp = uiApp ?? throw new ArgumentNullException(nameof(uiApp));
        }

        #endregion

        #region Public Methods

        public void Start()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(VoskVoiceCommandService));

            if (_isRunning)
                return;

            lock (_lock)
            {
                if (_isRunning) return;

                try
                {
                    // Suppress Vosk's internal logging
                    Vosk.Vosk.SetLogLevel(-1);

                    // Find model path
                    string modelPath = FindModelPath();
                    if (string.IsNullOrEmpty(modelPath) || !Directory.Exists(modelPath))
                    {
                        Log.Error("Vosk model not found. Expected in {PluginDir}/VoskModel/");
                        throw new DirectoryNotFoundException("Vosk model not found. Place the model in the VoskModel folder next to the plugin DLL.");
                    }

                    // Initialize Vosk model
                    _model = new Vosk.Model(modelPath);
                    Log.Debug("Vosk model loaded from: {Path}", modelPath);

                    // Build grammar from supported commands
                    string grammarJson = VoskVoiceCommandRouter.GetGrammar();
                    _recognizer = new Vosk.VoskRecognizer(_model, SampleRate, grammarJson);
                    Log.Debug("Vosk recognizer initialized with grammar");

                    // Initialize audio capture
                    _waveIn = new WaveInEvent
                    {
                        DeviceNumber = 0,
                        WaveFormat = new WaveFormat(SampleRate, 16, Channels)
                    };

                    _waveIn.DataAvailable += OnAudioDataAvailable;
                    _waveIn.RecordingStopped += OnRecordingStopped;

                    // Start recording
                    _waveIn.StartRecording();
                    _isRunning = true;

                    // Attach Revit Idling handler for thread-safe command execution
                    if (!_isHandlerAttached)
                    {
                        _uiApp.Idling += OnRevitIdling;
                        _isHandlerAttached = true;
                    }

                    Log.Information("Voice recognition service started successfully");
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "Failed to start voice recognition service");
                    Cleanup();
                    throw;
                }
            }
        }

        public void Stop()
        {
            if (!_isRunning) return;

            lock (_lock)
            {
                if (!_isRunning) return;

                try
                {
                    Cleanup();
                    Log.Information("Voice recognition service stopped");
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Error during voice service shutdown");
                }
            }
        }

        public void Toggle()
        {
            if (_isRunning)
                Stop();
            else
                Start();
        }

        #endregion

        #region Model Path Resolution

        /// <summary>
        /// Finds the Vosk model directory in {PluginDir}/VoskModel/.
        /// Returns null if not found.
        /// </summary>
        private string FindModelPath()
        {
            string pluginDir = Path.GetDirectoryName(typeof(VoskVoiceCommandService).Assembly.Location);
            if (!string.IsNullOrEmpty(pluginDir))
            {
                string modelPath = Path.Combine(pluginDir, "VoskModel");
                if (Directory.Exists(modelPath))
                {
                    // Look for a model subdirectory (e.g., vosk-model-small-en-us-0.15)
                    var modelDirs = Directory.GetDirectories(modelPath, "vosk-model*");
                    if (modelDirs.Length > 0)
                        return modelDirs[0];

                    // The directory itself may contain model files directly
                    if (File.Exists(Path.Combine(modelPath, "am", "final.mdl")) ||
                        File.Exists(Path.Combine(modelPath, "conf", "model.conf")))
                        return modelPath;
                }
            }

            return null;
        }

        #endregion

        #region Audio Processing

        private void OnAudioDataAvailable(object sender, WaveInEventArgs e)
        {
            lock (_lock)
            {
                if (_recognizer == null || !_isRunning)
                    return;

                try
                {
                    if (_recognizer.AcceptWaveform(e.Buffer, e.BytesRecorded))
                    {
                        string resultJson = _recognizer.Result();
                        string text = ExtractTextFromResult(resultJson);

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            if (!IsDuplicateCommand(text))
                            {
                                Log.Debug("Voice recognized: '{Text}'", text);
                                _commandQueue.Enqueue(text);
                                OnCommandRecognized?.Invoke(text);
                            }
                            else
                            {
                                Log.Debug("Duplicate command ignored: '{Text}'", text);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Error processing audio data");
                }
            }
        }

        private void OnRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                Log.Error(e.Exception, "Audio recording stopped due to error");
            }
            else
            {
                Log.Debug("Audio recording stopped");
            }
        }

        /// <summary>
        /// Extracts the "text" field from Vosk JSON result using System.Text.Json.
        /// </summary>
        private static string ExtractTextFromResult(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("text", out var textElement))
                {
                    return textElement.GetString()?.Trim().ToLowerInvariant();
                }
                return null;
            }
            catch
            {
                return null;
            }
        }

        private bool IsDuplicateCommand(string text)
        {
            var now = DateTime.UtcNow;

            if (text == _lastCommandText &&
                (now - _lastCommandTime).TotalMilliseconds < DuplicateThresholdMs)
            {
                return true;
            }

            _lastCommandText = text;
            _lastCommandTime = now;
            return false;
        }

        #endregion

        #region Revit Command Execution

        private void OnRevitIdling(object sender, IdlingEventArgs e)
        {
            while (_commandQueue.TryDequeue(out var text))
            {
                ExecuteCommand(text);
            }
        }

        private void ExecuteCommand(string text)
        {
            try
            {
                if (VoskVoiceCommandRouter.TryGetCommand(text, out var postableCmd, out var customAction))
                {
                    if (postableCmd.HasValue)
                    {
                        var cmdId = RevitCommandId.LookupPostableCommandId(postableCmd.Value);
                        _uiApp.PostCommand(cmdId);
                        Log.Information("Executed Revit command via voice: '{Command}' -> {PostableCommand}", text, postableCmd.Value);
                        OnCommandExecuted?.Invoke(text);
                    }
                    else if (customAction != null)
                    {
                        customAction.Invoke(_uiApp);
                        Log.Information("Executed custom command via voice: '{Command}'", text);
                        OnCommandExecuted?.Invoke(text);
                    }
                }
                else
                {
                    Log.Debug("Unrecognized voice command: '{Command}'", text);
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to execute voice command: '{Command}'", text);
            }
        }

        #endregion

        #region Cleanup & Disposal

        private void Cleanup()
        {
            _isRunning = false;

            if (_waveIn != null)
            {
                _waveIn.DataAvailable -= OnAudioDataAvailable;
                _waveIn.RecordingStopped -= OnRecordingStopped;

                try
                {
                    _waveIn.StopRecording();
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Error stopping recording");
                }

                _waveIn.Dispose();
                _waveIn = null;
            }

            if (_recognizer != null)
            {
                _recognizer.Dispose();
                _recognizer = null;
            }

            if (_model != null)
            {
                _model.Dispose();
                _model = null;
            }

            if (_isHandlerAttached)
            {
                try
                {
                    _uiApp.Idling -= OnRevitIdling;
                }
                catch (Exception ex)
                {
                    Log.Debug(ex, "Error detaching Idling handler");
                }
                _isHandlerAttached = false;
            }

            while (_commandQueue.TryDequeue(out _)) { }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                Stop();
            }

            _disposed = true;
        }

        ~VoskVoiceCommandService()
        {
            Dispose(false);
        }

        #endregion
    }
}
