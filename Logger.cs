using System;
using System.IO;

namespace BuildingBlocksManager
{
    public class Logger
    {
        private static readonly string LogDirectory = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "Logs");
        private readonly string logFile;

        public Logger()
        {
            Directory.CreateDirectory(LogDirectory);
            logFile = Path.Combine(LogDirectory, $"BBM_{DateTime.Now:yyyyMMdd_HHmmss}.log");
        }

        public void Info(string message)
        {
            WriteLog("INFO", message);
        }

        public void Warning(string message)
        {
            WriteLog("WARNING", message);
        }

        public void Error(string message)
        {
            WriteLog("ERROR", message);
        }

        public void Success(string message)
        {
            WriteLog("SUCCESS", message);
        }

        public void WriteExportErrorLog(int successCount, int errorCount, System.Collections.Generic.List<string> errors)
        {
            if (errorCount == 0) return;

            try
            {
                var errorLogFile = Path.Combine(LogDirectory, $"Export_Errors_{DateTime.Now:yyyyMMdd_HHmmss}.log");
                using (var writer = new StreamWriter(errorLogFile))
                {
                    writer.WriteLine($"Export Error Report - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    writer.WriteLine($"Successfully Exported: {successCount}");
                    writer.WriteLine($"Failed Exports: {errorCount}");
                    writer.WriteLine(new string('=', 50));
                    writer.WriteLine();

                    foreach (var error in errors)
                    {
                        writer.WriteLine(error);
                        writer.WriteLine();
                    }
                }

                Info($"Export error log created: {Path.GetFileName(errorLogFile)}");
            }
            catch (Exception ex)
            {
                Error($"Failed to create export error log: {ex.Message}");
            }
        }

        private void WriteLog(string level, string message)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {level}: {message}";
                File.AppendAllText(logFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        public static void CleanupOldLogs()
        {
            // TODO: Implement log cleanup
            // Remove logs older than 30 days
            try
            {
                if (Directory.Exists(LogDirectory))
                {
                    var files = Directory.GetFiles(LogDirectory, "BBM_*.log");
                    foreach (var file in files)
                    {
                        var fileInfo = new FileInfo(file);
                        if (fileInfo.CreationTime < DateTime.Now.AddDays(-30))
                        {
                            File.Delete(file);
                        }
                    }
                }
            }
            catch
            {
                // Silently handle cleanup errors
            }
        }
    }
}