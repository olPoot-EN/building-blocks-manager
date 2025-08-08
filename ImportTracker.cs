using System;
using System.Collections.Generic;
using System.IO;

namespace BuildingBlocksManager
{
    public class ImportTracker
    {
        private static readonly string TrackingDirectory = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager");
        private static readonly string TrackingFile = Path.Combine(TrackingDirectory, "import_tracking.txt");

        private Dictionary<string, DateTime> importHistory;

        public ImportTracker()
        {
            importHistory = new Dictionary<string, DateTime>();
            Load();
        }

        public DateTime GetLastImportTime(string filePath)
        {
            return importHistory.ContainsKey(filePath) ? importHistory[filePath] : DateTime.MinValue;
        }

        public void UpdateImportTime(string filePath, DateTime importTime)
        {
            importHistory[filePath] = importTime;
            Save();
        }

        public void UpdateImportTime(string filePath)
        {
            UpdateImportTime(filePath, DateTime.Now);
        }

        private void Load()
        {
            // TODO: Implement loading from tracking file
            // Format: filepath|timestamp
            try
            {
                if (File.Exists(TrackingFile))
                {
                    var lines = File.ReadAllLines(TrackingFile);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('|');
                        if (parts.Length == 2 && DateTime.TryParse(parts[1], out DateTime timestamp))
                        {
                            importHistory[parts[0]] = timestamp;
                        }
                    }
                }
            }
            catch
            {
                // Silently handle errors
            }
        }

        private void Save()
        {
            // TODO: Implement saving to tracking file
            try
            {
                Directory.CreateDirectory(TrackingDirectory);
                
                var lines = new List<string>();
                foreach (var kvp in importHistory)
                {
                    lines.Add($"{kvp.Key}|{kvp.Value:yyyy-MM-dd HH:mm}");
                }
                
                File.WriteAllLines(TrackingFile, lines);
            }
            catch
            {
                // Silently handle errors
            }
        }
    }
}