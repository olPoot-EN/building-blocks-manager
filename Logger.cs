using System;
using System.IO;
using System.Linq;

namespace BuildingBlocksManager
{
    public class Logger
    {
        private readonly string logDirectory;
        private readonly string generalLogFile;
        private readonly string exportLogFile;
        private readonly string errorLogFile;
        private readonly string deleteLogFile;
        private readonly bool enableDetailedLogging;
        private readonly string sessionId;

        public string GetLogDirectory() => logDirectory;

        public Logger(string templatePath = null, string sourceDirectoryPath = null, bool logToTemplateDirectory = true, bool enableDetailedLogging = true)
        {
            this.enableDetailedLogging = enableDetailedLogging;
            this.sessionId = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            
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
            
            // Define persistent log files
            generalLogFile = Path.Combine(logDirectory, "general.log");
            exportLogFile = Path.Combine(logDirectory, "export.log");
            errorLogFile = Path.Combine(logDirectory, "error.log");
            deleteLogFile = Path.Combine(logDirectory, "delete.log");
            
            // Write session start marker to general log
            WriteSessionMarker(generalLogFile, "SESSION START");
        }

        public void Info(string message)
        {
            if (enableDetailedLogging)
                WriteToFile(generalLogFile, "INFO", message);
        }

        public void Warning(string message)
        {
            WriteToFile(generalLogFile, "WARNING", message);
        }

        public void Error(string message)
        {
            WriteToFile(errorLogFile, "ERROR", message);
        }

        public void Success(string message)
        {
            WriteToFile(generalLogFile, "SUCCESS", message);
        }

        public void LogExport(string buildingBlockName, string category, string exportPath)
        {
            var message = $"{buildingBlockName} (Category: {category}) to {exportPath}";
            WriteToFile(exportLogFile, "EXPORT", message);
        }

        public void LogError(string operation, string buildingBlockName, string category, string errorMessage)
        {
            var message = $"{operation} FAILED: {buildingBlockName} (Category: {category}) - {errorMessage}";
            WriteToFile(errorLogFile, "ERROR", message);
        }

        public void LogDeletion(string buildingBlockName, string category)
        {
            var message = $"Deleted: {buildingBlockName} (Category: {category})";
            WriteToFile(deleteLogFile, "DELETE", message);
        }

        private void WriteToFile(string filePath, string level, string message)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {level}: {message}";
                File.AppendAllText(filePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        private void WriteSessionMarker(string filePath, string marker)
        {
            try
            {
                var sessionMarker = $"\n========== {marker}: {sessionId} ==========\n";
                File.AppendAllText(filePath, sessionMarker);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        public void EndSession()
        {
            WriteSessionMarker(generalLogFile, "SESSION END");
        }

        public void CleanupOldLogs()
        {
            // Archive and truncate logs that are too large (over 10MB) or very old content
            try
            {
                if (Directory.Exists(logDirectory))
                {
                    var logFiles = new[] { generalLogFile, exportLogFile, errorLogFile, deleteLogFile };
                    
                    foreach (var logFile in logFiles)
                    {
                        if (File.Exists(logFile))
                        {
                            var fileInfo = new FileInfo(logFile);
                            
                            // If file is larger than 10MB, archive it and start fresh
                            if (fileInfo.Length > 10 * 1024 * 1024)
                            {
                                var archiveName = Path.ChangeExtension(logFile, $".archive_{DateTime.Now:yyyyMMdd}.log");
                                File.Move(logFile, archiveName);
                                
                                // Create new file with session marker
                                WriteSessionMarker(logFile, "SESSION START (After Archive)");
                            }
                        }
                    }
                    
                    // Clean up old archive files (older than 90 days)
                    var archiveFiles = Directory.GetFiles(logDirectory, "*.archive_*.log");
                    foreach (var archiveFile in archiveFiles)
                    {
                        var fileInfo = new FileInfo(archiveFile);
                        if (fileInfo.CreationTime < DateTime.Now.AddDays(-90))
                        {
                            File.Delete(archiveFile);
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