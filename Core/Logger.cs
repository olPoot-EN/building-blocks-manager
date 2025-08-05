using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BuildingBlocksManager.Core
{
    public class Logger
    {
        private readonly string _logDirectory;
        private readonly string _currentLogFile;
        private static readonly object _lockObject = new object();

        public enum LogLevel
        {
            INFO,
            SUCCESS,
            WARNING,
            ERROR,
            TRACE
        }

        public class LogEntry
        {
            public DateTime Timestamp { get; set; }
            public LogLevel Level { get; set; }
            public string Message { get; set; }
            public string Category { get; set; }
        }

        public class FileExchangeInfo
        {
            public string Operation { get; set; } // "Import", "Export", "Scan"
            public string FilePath { get; set; }
            public string BuildingBlockName { get; set; }
            public string Category { get; set; }
            public bool Success { get; set; }
            public string ErrorMessage { get; set; }
            public DateTime Timestamp { get; set; }
        }

        public Logger()
        {
            _logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BuildingBlocksManager",
                "Logs"
            );

            Directory.CreateDirectory(_logDirectory);

            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            _currentLogFile = Path.Combine(_logDirectory, $"BBM_{timestamp}.log");

            // Log startup
            Log(LogLevel.INFO, "Building Blocks Manager started", "System");
        }

        public void Log(LogLevel level, string message, string category = "General")
        {
            var entry = new LogEntry
            {
                Timestamp = DateTime.Now,
                Level = level,
                Message = message,
                Category = category
            };

            WriteLogEntry(entry);
        }

        public void LogInfo(string message, string category = "General")
        {
            Log(LogLevel.INFO, message, category);
        }

        public void LogSuccess(string message, string category = "General")
        {
            Log(LogLevel.SUCCESS, message, category);
        }

        public void LogWarning(string message, string category = "General")
        {
            Log(LogLevel.WARNING, message, category);
        }

        public void LogError(string message, string category = "General")
        {
            Log(LogLevel.ERROR, message, category);
        }

        public void LogError(Exception ex, string context = "")
        {
            var message = string.IsNullOrEmpty(context) 
                ? $"Exception: {ex.Message}" 
                : $"{context}: {ex.Message}";
            
            Log(LogLevel.ERROR, message, "Exception");
            
            if (ex.InnerException != null)
            {
                Log(LogLevel.ERROR, $"Inner Exception: {ex.InnerException.Message}", "Exception");
            }
        }

        public void LogTrace(string message, string category = "Trace")
        {
            Log(LogLevel.TRACE, message, category);
        }

        public void LogFileExchange(FileExchangeInfo exchangeInfo)
        {
            var direction = exchangeInfo.Operation == "Import" ? "IN" : "OUT";
            var status = exchangeInfo.Success ? "SUCCESS" : "FAILED";
            
            var message = $"{exchangeInfo.Operation.ToUpper()} {direction}: {Path.GetFileName(exchangeInfo.FilePath)} -> BB:{exchangeInfo.BuildingBlockName} [{exchangeInfo.Category}] - {status}";
            
            if (!exchangeInfo.Success && !string.IsNullOrEmpty(exchangeInfo.ErrorMessage))
            {
                message += $" - {exchangeInfo.ErrorMessage}";
            }

            var level = exchangeInfo.Success ? LogLevel.SUCCESS : LogLevel.ERROR;
            Log(level, message, "FileExchange");
        }

        public void LogNewFilesDetected(List<string> newFiles, List<string> newBuildingBlockNames)
        {
            LogWarning($"NEW FILES DETECTED: {newFiles.Count} new autotext files found", "NewFiles");
            
            for (int i = 0; i < newFiles.Count; i++)
            {
                var fileName = Path.GetFileName(newFiles[i]);
                var bbName = i < newBuildingBlockNames.Count ? newBuildingBlockNames[i] : "Unknown";
                LogWarning($"  - {fileName} -> Building Block: '{bbName}'", "NewFiles");
            }

            LogWarning("IMPORTANT: New autotext entries will require Automator script updates for programmatic targeting", "AutomatorWarning");
            LogWarning("Please update your Automator configuration to reference these new Building Block names:", "AutomatorWarning");
            
            foreach (var bbName in newBuildingBlockNames)
            {
                LogWarning($"  - Reference in Automator: \"{bbName}\"", "AutomatorWarning");
            }
        }

        public void LogMissingFiles(List<string> missingFiles)
        {
            if (!missingFiles.Any())
                return;

            LogWarning($"MISSING FILES: {missingFiles.Count} previously present files are now missing", "MissingFiles");
            
            foreach (var missingFile in missingFiles)
            {
                var fileName = Path.GetFileName(missingFile);
                LogWarning($"  - Previously present: {fileName} (Full path: {missingFile})", "MissingFiles");
            }
        }

        public void LogRemovedFiles(List<string> removedFiles)
        {
            if (!removedFiles.Any())
                return;

            LogInfo($"REMOVED FILES: {removedFiles.Count} files were removed since last scan", "RemovedFiles");
            
            foreach (var removedFile in removedFiles)
            {
                var fileName = Path.GetFileName(removedFile);
                LogInfo($"  - Removed: {fileName}", "RemovedFiles");
            }
        }

        public void LogDirectoryScan(string directory, int totalFiles, int validFiles, int invalidFiles, int ignoredFiles, TimeSpan scanTime)
        {
            LogInfo($"Directory scan completed: {directory}", "DirectoryScan");
            LogInfo($"  - Total files scanned: {totalFiles}", "DirectoryScan");
            LogInfo($"  - Valid AT_ files: {validFiles}", "DirectoryScan");
            LogInfo($"  - Invalid files: {invalidFiles}", "DirectoryScan");
            LogInfo($"  - Ignored files: {ignoredFiles}", "DirectoryScan");
            LogInfo($"  - Scan duration: {scanTime.TotalSeconds:F2} seconds", "DirectoryScan");
        }

        public void LogImportSummary(int importedCount, int failedCount, int skippedCount, TimeSpan processingTime)
        {
            LogInfo($"Import operation completed", "ImportSummary");
            LogInfo($"  - Building Blocks imported: {importedCount}", "ImportSummary");
            LogInfo($"  - Failed imports: {failedCount}", "ImportSummary");
            LogInfo($"  - Skipped files: {skippedCount}", "ImportSummary");
            LogInfo($"  - Processing time: {processingTime.TotalSeconds:F2} seconds", "ImportSummary");
        }

        public void LogExportSummary(int exportedCount, int failedCount, TimeSpan processingTime, string exportPath)
        {
            LogInfo($"Export operation completed", "ExportSummary");
            LogInfo($"  - Building Blocks exported: {exportedCount}", "ExportSummary");
            LogInfo($"  - Failed exports: {failedCount}", "ExportSummary");
            LogInfo($"  - Export location: {exportPath}", "ExportSummary");
            LogInfo($"  - Processing time: {processingTime.TotalSeconds:F2} seconds", "ExportSummary");
        }

        public void LogBackupCreated(string templatePath, string backupPath)
        {
            LogInfo($"Template backup created: {Path.GetFileName(backupPath)}", "Backup");
            LogTrace($"  - Source: {templatePath}", "Backup");
            LogTrace($"  - Backup: {backupPath}", "Backup");
        }

        public void LogBackupRestored(string backupPath, string templatePath)
        {
            LogWarning($"Template restored from backup: {Path.GetFileName(backupPath)}", "Backup");
            LogInfo($"  - Restored to: {templatePath}", "Backup");
        }

        private void WriteLogEntry(LogEntry entry)
        {
            lock (_lockObject)
            {
                try
                {
                    var logLine = $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss}] {entry.Level}: {entry.Message}";
                    
                    if (!string.IsNullOrEmpty(entry.Category) && entry.Category != "General")
                    {
                        logLine = $"[{entry.Timestamp:yyyy-MM-dd HH:mm:ss}] {entry.Level} [{entry.Category}]: {entry.Message}";
                    }

                    File.AppendAllText(_currentLogFile, logLine + Environment.NewLine);
                }
                catch (Exception)
                {
                    // Silent fail - don't crash application if logging fails
                }
            }
        }

        public void CleanupOldLogs(int retentionDays = 30)
        {
            try
            {
                var cutoffDate = DateTime.Now.AddDays(-retentionDays);
                var logFiles = Directory.GetFiles(_logDirectory, "BBM_*.log");

                foreach (var logFile in logFiles)
                {
                    var fileInfo = new FileInfo(logFile);
                    if (fileInfo.CreationTime < cutoffDate)
                    {
                        try
                        {
                            File.Delete(logFile);
                            LogTrace($"Deleted old log file: {Path.GetFileName(logFile)}", "Cleanup");
                        }
                        catch (Exception)
                        {
                            // Ignore errors when deleting old logs
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex, "Failed to cleanup old log files");
            }
        }

        public List<LogEntry> ReadLogEntries(DateTime? fromDate = null, LogLevel? filterLevel = null)
        {
            var entries = new List<LogEntry>();
            
            try
            {
                if (!File.Exists(_currentLogFile))
                    return entries;

                var lines = File.ReadAllLines(_currentLogFile);
                
                foreach (var line in lines)
                {
                    var entry = ParseLogLine(line);
                    if (entry != null)
                    {
                        if (fromDate.HasValue && entry.Timestamp < fromDate.Value)
                            continue;

                        if (filterLevel.HasValue && entry.Level != filterLevel.Value)
                            continue;

                        entries.Add(entry);
                    }
                }
            }
            catch (Exception)
            {
                // Return empty list if reading fails
            }

            return entries;
        }

        private LogEntry ParseLogLine(string line)
        {
            try
            {
                // Parse format: [2025-08-03 14:30:15] INFO [Category]: Message
                // or: [2025-08-03 14:30:15] INFO: Message
                
                if (!line.StartsWith("["))
                    return null;

                var timestampEnd = line.IndexOf("] ");
                if (timestampEnd == -1)
                    return null;

                var timestampStr = line.Substring(1, timestampEnd - 1);
                if (!DateTime.TryParseExact(timestampStr, "yyyy-MM-dd HH:mm:ss", 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, out DateTime timestamp))
                    return null;

                var remainder = line.Substring(timestampEnd + 2);
                var levelEnd = remainder.IndexOf(":");
                if (levelEnd == -1)
                    return null;

                var levelPart = remainder.Substring(0, levelEnd).Trim();
                var messagePart = remainder.Substring(levelEnd + 1).Trim();

                string category = "General";
                string levelStr = levelPart;

                // Check if there's a category in brackets
                if (levelPart.Contains("[") && levelPart.Contains("]"))
                {
                    var categoryStart = levelPart.IndexOf("[");
                    var categoryEnd = levelPart.IndexOf("]");
                    category = levelPart.Substring(categoryStart + 1, categoryEnd - categoryStart - 1);
                    levelStr = levelPart.Substring(0, categoryStart).Trim();
                }

                if (!Enum.TryParse<LogLevel>(levelStr, out LogLevel level))
                    return null;

                return new LogEntry
                {
                    Timestamp = timestamp,
                    Level = level,
                    Category = category,
                    Message = messagePart
                };
            }
            catch
            {
                return null;
            }
        }

        public string GetCurrentLogFilePath()
        {
            return _currentLogFile;
        }

        public string GetLogDirectory()
        {
            return _logDirectory;
        }

        public void Dispose()
        {
            LogInfo("Building Blocks Manager session ended", "System");
        }
    }
}