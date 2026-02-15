using System;
using System.IO;
using Serilog;

namespace BimTasksV2.Logging
{
    /// <summary>
    /// Configures Serilog file logging for BimTasksV2.
    /// Logs are written to %APPDATA%/BimTasks/Logs/ with daily rolling files.
    /// </summary>
    public static class SerilogConfig
    {
        public static void Initialize()
        {
            string logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "BimTasks",
                "Logs");

            // Ensure log directory exists
            Directory.CreateDirectory(logDirectory);

            string logPath = Path.Combine(logDirectory, "bimtasks-.log");

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    logPath,
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            Log.Information("BimTasksV2 Serilog initialized. Log path: {LogPath}", logPath);
        }
    }
}
