using System;
using System.IO;

namespace BuildingBlocksManager
{
    public class Logger
    {
        private readonly string logDirectory;
        private readonly string logFile;

        public string GetLogDirectory() => logDirectory;

        public Logger(string sourceDirectoryPath = null)
        {
            if (!string.IsNullOrEmpty(sourceDirectoryPath) && Directory.Exists(sourceDirectoryPath))
            {
                logDirectory = Path.Combine(sourceDirectoryPath, "Logs");
            }
            else
            {
                // Fallback to user's local app data if source directory not provided or doesn't exist
                logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "Logs");
            }
            
            Directory.CreateDirectory(logDirectory);
            logFile = Path.Combine(logDirectory, $"BBM_{DateTime.Now:yyyyMMdd_HHmmss}.log");
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

        public void LogExport(string buildingBlockName, string category, string exportPath)
        {
            try
            {
                var exportLogFile = Path.Combine(logDirectory, "Export.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Exported: {buildingBlockName} (Category: {category}) to {exportPath}";
                File.AppendAllText(exportLogFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        public void LogError(string operation, string buildingBlockName, string category, string errorMessage)
        {
            try
            {
                var errorLogFile = Path.Combine(logDirectory, "Error.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {operation} FAILED: {buildingBlockName} (Category: {category}) - {errorMessage}";
                File.AppendAllText(errorLogFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        public void LogDeletion(string buildingBlockName, string category)
        {
            try
            {
                var deleteLogFile = Path.Combine(logDirectory, "Delete.log");
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Deleted: {buildingBlockName} (Category: {category})";
                File.AppendAllText(deleteLogFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
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

        public void CleanupOldLogs()
        {
            // Remove logs older than 30 days
            try
            {
                if (Directory.Exists(logDirectory))
                {
                    var files = Directory.GetFiles(logDirectory, "BBM_*.log");
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