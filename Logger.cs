using System;
using System.IO;
using System.Linq;

namespace BuildingBlocksManager
{
    public class Logger
    {
        private readonly string logDirectory;
        private readonly string sessionDirectory;
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
            this.sessionId = DateTime.Now.ToString("yyyy-MM-dd_HHmm");
            this.templatePath = templatePath;
            this.sourceDirectoryPath = sourceDirectoryPath;
            
            // Determine log directory based on settings with proper fallback handling
            string primaryLogDirectory = null;
            
            if (logToTemplateDirectory && !string.IsNullOrEmpty(templatePath) && File.Exists(templatePath))
            {
                // Use template directory for logs
                var templateDir = Path.GetDirectoryName(templatePath);
                primaryLogDirectory = Path.Combine(templateDir, "BBM_Logs");
            }
            else if (!string.IsNullOrEmpty(sourceDirectoryPath) && Directory.Exists(sourceDirectoryPath))
            {
                // Fall back to source directory
                primaryLogDirectory = Path.Combine(sourceDirectoryPath, "BBM_Logs");
            }
            
            // Try to create primary log directory, fall back to local app data if fails
            logDirectory = TryCreateLogDirectory(primaryLogDirectory) ?? 
                          TryCreateLogDirectory(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "BBM_Logs")) ??
                          Path.GetTempPath(); // Final fallback to temp directory
            
            // Create session-specific subdirectory with error handling
            sessionDirectory = Path.Combine(logDirectory, sessionId);
            try
            {
                Directory.CreateDirectory(sessionDirectory);
            }
            catch
            {
                // If session directory creation fails, use the log directory itself
                sessionDirectory = logDirectory;
            }
            
            // Define session-specific log files
            generalLogFile = Path.Combine(sessionDirectory, "general.log");
            importLogFile = Path.Combine(sessionDirectory, "import.log");
            exportLogFile = Path.Combine(sessionDirectory, "export.log");
            errorLogFile = Path.Combine(sessionDirectory, "error.log");
            deleteLogFile = Path.Combine(sessionDirectory, "delete.log");
            
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
            
            // Create columnated delete log entry similar to ledger format
            try
            {
                // Check if file is new/empty and add header
                if (!File.Exists(deleteLogFile) || new FileInfo(deleteLogFile).Length == 0)
                {
                    var header = "# Building Block Deletions Log" + Environment.NewLine +
                                $"# Session: {sessionId}" + Environment.NewLine +
                                "# Date/Time".PadRight(20) + "\t" + "Template".PadRight(40) + "\t" + "Name".PadRight(50) + "\t" + "Category" + Environment.NewLine;
                    File.AppendAllText(deleteLogFile, header);
                }
                
                // Create columnated entry
                var dateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                var logEntry = $"{dateTime.PadRight(20)}\t{templateName.PadRight(40)}\t{buildingBlockName.PadRight(50)}\t{category ?? ""}";
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

        private string TryCreateLogDirectory(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath))
                return null;
                
            try
            {
                Directory.CreateDirectory(directoryPath);
                return directoryPath;
            }
            catch (UnauthorizedAccessException)
            {
                // Access denied - return null to try fallback
                return null;
            }
            catch (IOException)
            {
                // General IO issues - return null to try fallback
                return null;
            }
            catch (NotSupportedException)
            {
                // Path format issues - return null to try fallback
                return null;
            }
            catch
            {
                // Any other exception - return null to try fallback
                return null;
            }
        }

        public void CleanupOldLogs()
        {
            // Keep only the 10 most recent session folders
            try
            {
                if (Directory.Exists(logDirectory))
                {
                    var sessionFolders = Directory.GetDirectories(logDirectory)
                        .Where(dir => {
                            var folderName = Path.GetFileName(dir);
                            // Match pattern YYYY-MM-DD_HHMM
                            return folderName.Length == 16 && 
                                   folderName[4] == '-' && 
                                   folderName[7] == '-' && 
                                   folderName[10] == '_';
                        })
                        .Select(dir => new DirectoryInfo(dir))
                        .OrderByDescending(dir => dir.CreationTime)
                        .ToList();
                    
                    // Keep only the 10 most recent, delete the rest
                    var foldersToDelete = sessionFolders.Skip(10);
                    
                    foreach (var sessionFolder in foldersToDelete)
                    {
                        try
                        {
                            Directory.Delete(sessionFolder.FullName, true);
                        }
                        catch
                        {
                            // Silently handle errors with individual session folders
                        }
                    }
                    
                    // For the recent folders we're keeping, check if any individual log files are too large (over 10MB)
                    var foldersToKeep = sessionFolders.Take(10);
                    foreach (var sessionFolder in foldersToKeep)
                    {
                        try
                        {
                            var logFiles = Directory.GetFiles(sessionFolder.FullName, "*.log");
                            foreach (var logFile in logFiles)
                            {
                                var fileInfo = new FileInfo(logFile);
                                if (fileInfo.Length > 10 * 1024 * 1024)
                                {
                                    // Archive very large log files within the session folder
                                    var archiveName = Path.ChangeExtension(logFile, $".archive_{DateTime.Now:yyyyMMdd}.log");
                                    File.Move(logFile, archiveName);
                                }
                            }
                        }
                        catch
                        {
                            // Silently handle errors with individual session folders
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