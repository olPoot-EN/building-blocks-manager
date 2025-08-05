using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace BuildingBlocksManager.Core
{
    public class SettingsManager
    {
        private readonly string _settingsFilePath;
        private Settings _settings;

        public class Settings
        {
            public string LastSourceDirectory { get; set; } = "";
            public string LastTemplateFile { get; set; } = "";
            public string LastExportDirectory { get; set; } = "";
            public bool FlatImportEnabled { get; set; } = false;
            public bool FlatExportEnabled { get; set; } = false;
            public string DefaultFlatImportCategory { get; set; } = "InternalAutotext";
            public bool ShowWarningsForNewFiles { get; set; } = true;
            public bool ImportOnlyChanged { get; set; } = true;
            public List<string> RecentSourceDirectories { get; set; } = new List<string>();
            public List<string> RecentTemplateFiles { get; set; } = new List<string>();
            public List<string> RecentExportDirectories { get; set; } = new List<string>();
            public DateTime LastUsed { get; set; } = DateTime.Now;
            public string WindowSize { get; set; } = "800,600";
            public string WindowLocation { get; set; } = "";
            public bool EnableLogging { get; set; } = true;
            public int LogRetentionDays { get; set; } = 30;
        }

        public SettingsManager()
        {
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BuildingBlocksManager"
            );

            Directory.CreateDirectory(appDataPath);
            _settingsFilePath = Path.Combine(appDataPath, "settings.json");
            
            LoadSettings();
        }

        public Settings GetSettings()
        {
            return _settings;
        }

        public void UpdateLastSourceDirectory(string directory)
        {
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                return;

            _settings.LastSourceDirectory = directory;
            AddToRecentList(_settings.RecentSourceDirectories, directory, 10);
            SaveSettings();
        }

        public void UpdateLastTemplateFile(string templateFile)
        {
            if (string.IsNullOrEmpty(templateFile) || !File.Exists(templateFile))
                return;

            _settings.LastTemplateFile = templateFile;
            AddToRecentList(_settings.RecentTemplateFiles, templateFile, 10);
            SaveSettings();
        }

        public void UpdateLastExportDirectory(string directory)
        {
            if (string.IsNullOrEmpty(directory))
                return;

            _settings.LastExportDirectory = directory;
            AddToRecentList(_settings.RecentExportDirectories, directory, 10);
            SaveSettings();
        }

        public void UpdateImportOptions(bool flatImport, string flatImportCategory, bool showWarnings, bool importOnlyChanged)
        {
            _settings.FlatImportEnabled = flatImport;
            _settings.DefaultFlatImportCategory = flatImportCategory ?? "InternalAutotext";
            _settings.ShowWarningsForNewFiles = showWarnings;
            _settings.ImportOnlyChanged = importOnlyChanged;
            SaveSettings();
        }

        public void UpdateExportOptions(bool flatExport)
        {
            _settings.FlatExportEnabled = flatExport;
            SaveSettings();
        }

        public void UpdateWindowSettings(string size, string location)
        {
            if (!string.IsNullOrEmpty(size))
                _settings.WindowSize = size;
            
            if (!string.IsNullOrEmpty(location))
                _settings.WindowLocation = location;
            
            SaveSettings();
        }

        public void UpdateLoggingSettings(bool enableLogging, int retentionDays)
        {
            _settings.EnableLogging = enableLogging;
            _settings.LogRetentionDays = Math.Max(1, Math.Min(365, retentionDays)); // Clamp between 1-365 days
            SaveSettings();
        }

        public List<string> GetRecentSourceDirectories()
        {
            return new List<string>(_settings.RecentSourceDirectories);
        }

        public List<string> GetRecentTemplateFiles()
        {
            return new List<string>(_settings.RecentTemplateFiles);
        }

        public List<string> GetRecentExportDirectories()
        {
            return new List<string>(_settings.RecentExportDirectories);
        }

        public void ClearRecentLists()
        {
            _settings.RecentSourceDirectories.Clear();
            _settings.RecentTemplateFiles.Clear();
            _settings.RecentExportDirectories.Clear();
            SaveSettings();
        }

        public void RemoveFromRecentSources(string directory)
        {
            _settings.RecentSourceDirectories.RemoveAll(d => 
                string.Equals(d, directory, StringComparison.OrdinalIgnoreCase));
            
            if (string.Equals(_settings.LastSourceDirectory, directory, StringComparison.OrdinalIgnoreCase))
            {
                _settings.LastSourceDirectory = _settings.RecentSourceDirectories.Count > 0 
                    ? _settings.RecentSourceDirectories[0] 
                    : "";
            }
            
            SaveSettings();
        }

        public void RemoveFromRecentTemplates(string templateFile)
        {
            _settings.RecentTemplateFiles.RemoveAll(f => 
                string.Equals(f, templateFile, StringComparison.OrdinalIgnoreCase));
            
            if (string.Equals(_settings.LastTemplateFile, templateFile, StringComparison.OrdinalIgnoreCase))
            {
                _settings.LastTemplateFile = _settings.RecentTemplateFiles.Count > 0 
                    ? _settings.RecentTemplateFiles[0] 
                    : "";
            }
            
            SaveSettings();
        }

        public void CleanupNonExistentPaths()
        {
            // Clean up source directories
            var validSources = new List<string>();
            foreach (var dir in _settings.RecentSourceDirectories)
            {
                if (Directory.Exists(dir))
                    validSources.Add(dir);
            }
            _settings.RecentSourceDirectories = validSources;

            // Clean up template files
            var validTemplates = new List<string>();
            foreach (var file in _settings.RecentTemplateFiles)
            {
                if (File.Exists(file))
                    validTemplates.Add(file);
            }
            _settings.RecentTemplateFiles = validTemplates;

            // Update last used paths if they no longer exist
            if (!string.IsNullOrEmpty(_settings.LastSourceDirectory) && !Directory.Exists(_settings.LastSourceDirectory))
            {
                _settings.LastSourceDirectory = validSources.Count > 0 ? validSources[0] : "";
            }

            if (!string.IsNullOrEmpty(_settings.LastTemplateFile) && !File.Exists(_settings.LastTemplateFile))
            {
                _settings.LastTemplateFile = validTemplates.Count > 0 ? validTemplates[0] : "";
            }

            SaveSettings();
        }

        public void ResetToDefaults()
        {
            _settings = new Settings();
            SaveSettings();
        }

        public string GetSettingsFilePath()
        {
            return _settingsFilePath;
        }

        public Dictionary<string, object> GetSettingsSummary()
        {
            return new Dictionary<string, object>
            {
                ["LastSourceDirectory"] = _settings.LastSourceDirectory,
                ["LastTemplateFile"] = _settings.LastTemplateFile,
                ["LastExportDirectory"] = _settings.LastExportDirectory,
                ["FlatImportEnabled"] = _settings.FlatImportEnabled,
                ["FlatExportEnabled"] = _settings.FlatExportEnabled,
                ["ShowWarningsForNewFiles"] = _settings.ShowWarningsForNewFiles,
                ["ImportOnlyChanged"] = _settings.ImportOnlyChanged,
                ["RecentSourcesCount"] = _settings.RecentSourceDirectories.Count,
                ["RecentTemplatesCount"] = _settings.RecentTemplateFiles.Count,
                ["LogRetentionDays"] = _settings.LogRetentionDays,
                ["LastUsed"] = _settings.LastUsed,
                ["SettingsFile"] = _settingsFilePath
            };
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    var json = File.ReadAllText(_settingsFilePath);
                    _settings = JsonSerializer.Deserialize<Settings>(json) ?? new Settings();
                }
                else
                {
                    _settings = new Settings();
                }

                // Update last used time
                _settings.LastUsed = DateTime.Now;
            }
            catch (Exception)
            {
                // If settings file is corrupted, create new default settings
                _settings = new Settings();
            }
        }

        private void SaveSettings()
        {
            try
            {
                _settings.LastUsed = DateTime.Now;
                
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_settings, options);
                File.WriteAllText(_settingsFilePath, json);
            }
            catch (Exception)
            {
                // Silent fail - don't crash application if settings save fails
            }
        }

        private void AddToRecentList(List<string> list, string item, int maxItems)
        {
            // Remove if already exists
            list.RemoveAll(x => string.Equals(x, item, StringComparison.OrdinalIgnoreCase));
            
            // Add to beginning
            list.Insert(0, item);
            
            // Trim to max items
            if (list.Count > maxItems)
            {
                list.RemoveRange(maxItems, list.Count - maxItems);
            }
        }
    }
}