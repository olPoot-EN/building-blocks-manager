using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;

namespace BuildingBlocksManager.Core
{
    public class ImportTracker
    {
        private readonly string _trackingFilePath;
        private readonly string _manifestFilePath;
        private Dictionary<string, DateTime> _importedFiles;
        private HashSet<string> _lastKnownFiles;

        public class ChangeAnalysis
        {
            public List<string> NewFiles { get; set; } = new List<string>();
            public List<string> ModifiedFiles { get; set; } = new List<string>();
            public List<string> UpToDateFiles { get; set; } = new List<string>();
            public List<string> MissingFiles { get; set; } = new List<string>();
            public List<string> RemovedFiles { get; set; } = new List<string>();
        }

        public class FileChangeInfo
        {
            public string FilePath { get; set; }
            public DateTime CurrentModified { get; set; }
            public DateTime? LastImported { get; set; }
            public string Status { get; set; } // "New", "Modified", "UpToDate", "Missing"
            public string BuildingBlockName { get; set; }
        }

        public ImportTracker()
        {
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BuildingBlocksManager"
            );

            Directory.CreateDirectory(appDataPath);
            
            _trackingFilePath = Path.Combine(appDataPath, "import_tracking.txt");
            _manifestFilePath = Path.Combine(appDataPath, "file_manifest.txt");
            
            _importedFiles = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            _lastKnownFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            
            LoadTrackingData();
            LoadManifest();
        }

        private void LoadTrackingData()
        {
            if (!File.Exists(_trackingFilePath))
                return;

            try
            {
                var lines = File.ReadAllLines(_trackingFilePath);
                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var parts = line.Split('|');
                    if (parts.Length >= 2)
                    {
                        var filePath = parts[0].Trim();
                        if (DateTime.TryParseExact(parts[1].Trim(), "yyyy-MM-dd HH:mm:ss", 
                            CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime importTime))
                        {
                            _importedFiles[filePath] = importTime;
                        }
                    }
                }
            }
            catch (Exception)
            {
                // If tracking file is corrupted, start fresh
                _importedFiles.Clear();
            }
        }

        private void LoadManifest()
        {
            if (!File.Exists(_manifestFilePath))
                return;

            try
            {
                var lines = File.ReadAllLines(_manifestFilePath);
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        _lastKnownFiles.Add(line.Trim());
                    }
                }
            }
            catch (Exception)
            {
                // If manifest file is corrupted, start fresh
                _lastKnownFiles.Clear();
            }
        }

        public void UpdateImportTime(string filePath, DateTime importTime)
        {
            _importedFiles[filePath] = importTime;
            SaveTrackingData();
        }

        public void UpdateImportTime(string filePath)
        {
            UpdateImportTime(filePath, DateTime.Now);
        }

        public DateTime? GetLastImportTime(string filePath)
        {
            if (_importedFiles.TryGetValue(filePath, out DateTime importTime))
            {
                return importTime;
            }
            return null;
        }

        public bool HasBeenImported(string filePath)
        {
            return _importedFiles.ContainsKey(filePath);
        }

        public bool HasBeenModifiedSinceImport(string filePath)
        {
            if (!File.Exists(filePath))
                return false;

            var lastImported = GetLastImportTime(filePath);
            if (!lastImported.HasValue)
                return true; // Never imported, so consider it modified

            var fileModified = File.GetLastWriteTime(filePath);
            return fileModified > lastImported.Value;
        }

        public ChangeAnalysis AnalyzeChanges(IEnumerable<FileManager.FileInfo> currentFiles)
        {
            var analysis = new ChangeAnalysis();
            var currentFilePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var fileInfo in currentFiles.Where(f => f.IsValid))
            {
                currentFilePaths.Add(fileInfo.FullPath);
                
                if (!HasBeenImported(fileInfo.FullPath))
                {
                    analysis.NewFiles.Add(fileInfo.FullPath);
                }
                else if (HasBeenModifiedSinceImport(fileInfo.FullPath))
                {
                    analysis.ModifiedFiles.Add(fileInfo.FullPath);
                }
                else
                {
                    analysis.UpToDateFiles.Add(fileInfo.FullPath);
                }
            }

            // Find files that were previously known but are now missing
            foreach (var previousFile in _lastKnownFiles)
            {
                if (!currentFilePaths.Contains(previousFile))
                {
                    if (HasBeenImported(previousFile))
                    {
                        analysis.MissingFiles.Add(previousFile);
                    }
                    else
                    {
                        analysis.RemovedFiles.Add(previousFile);
                    }
                }
            }

            return analysis;
        }

        public List<FileChangeInfo> GetDetailedChangeInfo(IEnumerable<FileManager.FileInfo> currentFiles)
        {
            var changes = new List<FileChangeInfo>();

            foreach (var fileInfo in currentFiles.Where(f => f.IsValid))
            {
                var changeInfo = new FileChangeInfo
                {
                    FilePath = fileInfo.FullPath,
                    CurrentModified = fileInfo.LastModified,
                    LastImported = GetLastImportTime(fileInfo.FullPath),
                    BuildingBlockName = fileInfo.BuildingBlockName
                };

                if (!HasBeenImported(fileInfo.FullPath))
                {
                    changeInfo.Status = "New";
                }
                else if (HasBeenModifiedSinceImport(fileInfo.FullPath))
                {
                    changeInfo.Status = "Modified";
                }
                else
                {
                    changeInfo.Status = "UpToDate";
                }

                changes.Add(changeInfo);
            }

            // Add missing files
            foreach (var missingFile in _lastKnownFiles)
            {
                if (!currentFiles.Any(f => string.Equals(f.FullPath, missingFile, StringComparison.OrdinalIgnoreCase)))
                {
                    changes.Add(new FileChangeInfo
                    {
                        FilePath = missingFile,
                        LastImported = GetLastImportTime(missingFile),
                        Status = "Missing",
                        BuildingBlockName = ExtractBuildingBlockNameFromPath(missingFile)
                    });
                }
            }

            return changes.OrderBy(c => c.Status).ThenBy(c => c.FilePath).ToList();
        }

        private string ExtractBuildingBlockNameFromPath(string filePath)
        {
            try
            {
                var fileName = Path.GetFileName(filePath);
                if (fileName.StartsWith("AT_", StringComparison.OrdinalIgnoreCase))
                {
                    var nameWithoutPrefix = fileName.Substring(3);
                    return Path.GetFileNameWithoutExtension(nameWithoutPrefix).Replace(' ', '_');
                }
            }
            catch
            {
                // If extraction fails, return the filename
            }
            
            return Path.GetFileNameWithoutExtension(filePath);
        }

        public void UpdateManifest(IEnumerable<FileManager.FileInfo> currentFiles)
        {
            _lastKnownFiles.Clear();
            
            foreach (var fileInfo in currentFiles.Where(f => f.IsValid))
            {
                _lastKnownFiles.Add(fileInfo.FullPath);
            }

            SaveManifest();
        }

        public void RemoveFromTracking(string filePath)
        {
            _importedFiles.Remove(filePath);
            _lastKnownFiles.Remove(filePath);
            SaveTrackingData();
            SaveManifest();
        }

        public void ClearTrackingData()
        {
            _importedFiles.Clear();
            _lastKnownFiles.Clear();
            SaveTrackingData();
            SaveManifest();
        }

        public Dictionary<string, DateTime> GetAllImportedFiles()
        {
            return new Dictionary<string, DateTime>(_importedFiles);
        }

        public int GetImportedFileCount()
        {
            return _importedFiles.Count;
        }

        public DateTime? GetLastImportDate()
        {
            if (_importedFiles.Values.Any())
            {
                return _importedFiles.Values.Max();
            }
            return null;
        }

        private void SaveTrackingData()
        {
            try
            {
                var lines = _importedFiles.Select(kvp => 
                    $"{kvp.Key}|{kvp.Value:yyyy-MM-dd HH:mm:ss}");
                
                File.WriteAllLines(_trackingFilePath, lines);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save tracking data: {ex.Message}", ex);
            }
        }

        private void SaveManifest()
        {
            try
            {
                File.WriteAllLines(_manifestFilePath, _lastKnownFiles.OrderBy(f => f));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save manifest: {ex.Message}", ex);
            }
        }

        public string GetTrackingDataPath()
        {
            return _trackingFilePath;
        }

        public string GetManifestPath()
        {
            return _manifestFilePath;
        }

        public void CleanupOrphanedEntries(IEnumerable<FileManager.FileInfo> currentFiles)
        {
            var currentFilePaths = new HashSet<string>(
                currentFiles.Where(f => f.IsValid).Select(f => f.FullPath),
                StringComparer.OrdinalIgnoreCase
            );

            var orphanedEntries = _importedFiles.Keys
                .Where(path => !currentFilePaths.Contains(path) && !File.Exists(path))
                .ToList();

            foreach (var orphaned in orphanedEntries)
            {
                _importedFiles.Remove(orphaned);
                _lastKnownFiles.Remove(orphaned);
            }

            if (orphanedEntries.Any())
            {
                SaveTrackingData();
                SaveManifest();
            }
        }
    }
}