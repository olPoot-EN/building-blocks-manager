using System;
using System.IO;
using System.Linq;

namespace BuildingBlocksManager
{
    public class Logger
    {
        private readonly string logDirectory;
        private readonly string generalLogFile;
        private readonly string importLogFile;
        private readonly string exportLogFile;
        private readonly string errorLogFile;
        private readonly string deleteLogFile;
        private readonly bool enableDetailedLogging;
        private readonly string sessionId;
        private readonly string templatePath;
        private readonly string sourceDirectoryPath;

        public string GetLogDirectory() => logDirectory;

        public Logger(string templatePath = null, string sourceDirectoryPath = null, bool logToTemplateDirectory = true, bool enableDetailedLogging = true)
        {
            this.enableDetailedLogging = enableDetailedLogging;
            this.sessionId = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            this.templatePath = templatePath;
            this.sourceDirectoryPath = sourceDirectoryPath;
            
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
            importLogFile = Path.Combine(logDirectory, "import.log");
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

        public void StartImportSession()
        {
            WriteSessionHeader(importLogFile, "IMPORT SESSION START");
            WriteSessionInfo(importLogFile, "Source", sourceDirectoryPath);
            WriteSessionInfo(importLogFile, "Template", GetTemplateFileName());
        }

        public void EndImportSession(int itemCount)
        {
            WriteSessionFooter(importLogFile, "IMPORT SESSION END", $"{itemCount} items imported");
        }

        public void LogImport(string buildingBlockName, string category)
        {
            var message = $"    {buildingBlockName} ({category})"; // Indented for visual hierarchy
            WriteToFile(importLogFile, "", message); // No level prefix for session items
        }

        public void StartExportSession(string exportPath)
        {
            WriteSessionHeader(exportLogFile, "EXPORT SESSION START");
            WriteSessionInfo(exportLogFile, "Template", GetTemplateFileName());
            WriteSessionInfo(exportLogFile, "Destination", exportPath);
        }

        public void EndExportSession(int itemCount)
        {
            WriteSessionFooter(exportLogFile, "EXPORT SESSION END", $"{itemCount} items exported");
        }

        public void LogExport(string buildingBlockName, string category)
        {
            var message = $"    {buildingBlockName} ({category})"; // Indented for visual hierarchy
            WriteToFile(exportLogFile, "", message); // No level prefix for session items
        }

        public void LogError(string operation, string buildingBlockName, string category, string errorMessage)
        {
            var message = $"{operation} FAILED: {buildingBlockName} (Category: {category}) - {errorMessage}";
            WriteToFile(errorLogFile, "ERROR", message);
        }

        public void LogDeletion(string buildingBlockName, string category)
        {
            var templateName = GetTemplateFileName();
            var message = $"From: {templateName} --- {buildingBlockName} ({category})";
            // Use timestamp but no level for delete log
            try
            {
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm}] {message}";
                File.AppendAllText(deleteLogFile, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        private void WriteToFile(string filePath, string level, string message)
        {
            try
            {
                var logEntry = string.IsNullOrEmpty(level) 
                    ? message // Session items without timestamp or level
                    : $"[{DateTime.Now:yyyy-MM-dd HH:mm}] {level}: {message}";
                File.AppendAllText(filePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        private void WriteSessionHeader(string filePath, string sessionType)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm}] === {sessionType} ===";
                File.AppendAllText(filePath, logEntry + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        private void WriteSessionInfo(string filePath, string label, string value)
        {
            try
            {
                if (!string.IsNullOrEmpty(value))
                {
                    var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm}] {label}: {value}";
                    File.AppendAllText(filePath, logEntry + Environment.NewLine);
                }
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        private void WriteSessionFooter(string filePath, string sessionType, string summary)
        {
            try
            {
                var logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm}] === {sessionType} === ({summary})";
                File.AppendAllText(filePath, logEntry + Environment.NewLine + Environment.NewLine);
            }
            catch
            {
                // Silently handle logging errors
            }
        }

        private string GetTemplateFileName()
        {
            return !string.IsNullOrEmpty(templatePath) ? Path.GetFileName(templatePath) : "Unknown";
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
                    var logFiles = new[] { generalLogFile, importLogFile, exportLogFile, errorLogFile, deleteLogFile };
                    
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