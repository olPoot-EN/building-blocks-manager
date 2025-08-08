using System;
using System.IO;
using System.Linq;

namespace BuildingBlocksManager
{
    public class Logger
    {
        private readonly string logDirectory;
        private readonly string logFile;
        private readonly bool enableDetailedLogging;

        public string GetLogDirectory() => logDirectory;

        public Logger(string templatePath = null, string sourceDirectoryPath = null, bool logToTemplateDirectory = true, bool enableDetailedLogging = true)
        {
            this.enableDetailedLogging = enableDetailedLogging;
            
            // Determine log directory based on settings
            if (logToTemplateDirectory && !string.IsNullOrEmpty(templatePath) && File.Exists(templatePath))
            {
                // Use template directory for logs
                var templateDir = Path.GetDirectoryName(templatePath);
                logDirectory = Path.Combine(templateDir, "BBM_Logs");
            }
            else if (!string.IsNullOrEmpty(sourceDirectoryPath) && Directory.Exists(sourceDirectoryPath))
            {
                // Fall back to source directory
                logDirectory = Path.Combine(sourceDirectoryPath, "BBM_Logs");
            }
            else
            {
                // Final fallback to user's local app data
                logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "BBM_Logs");
            }
            
            Directory.CreateDirectory(logDirectory);
            
            // Use a more readable filename
            var sessionTime = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            logFile = Path.Combine(logDirectory, $"BBM_Session_{sessionTime}.log");
        }

        public void Info(string message)
        {
            if (enableDetailedLogging)
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
                    var files = Directory.GetFiles(logDirectory, "BBM_Session_*.log");
                    var exportFiles = Directory.GetFiles(logDirectory, "Export.log");
                    var errorFiles = Directory.GetFiles(logDirectory, "Error.log");
                    var deleteFiles = Directory.GetFiles(logDirectory, "Delete.log");
                    
                    var allFiles = files.Concat(exportFiles).Concat(errorFiles).Concat(deleteFiles);
                    
                    foreach (var file in allFiles)
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