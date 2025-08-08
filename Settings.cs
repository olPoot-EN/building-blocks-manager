using System;
using System.IO;

namespace BuildingBlocksManager
{
    public class Settings
    {
        private static readonly string SettingsDirectory = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager");
        private static readonly string SettingsFile = Path.Combine(SettingsDirectory, "settings.txt");

        public string LastTemplatePath { get; set; } = "";
        public string LastSourceDirectory { get; set; } = "";
        public string LastExportDirectory { get; set; } = "";
        public bool FlatImport { get; set; } = false;
        public bool FlatExport { get; set; } = false;
        public bool LogToTemplateDirectory { get; set; } = true;
        public bool EnableDetailedLogging { get; set; } = true;

        public static Settings Load()
        {
            var settings = new Settings();
            
            try
            {
                if (File.Exists(SettingsFile))
                {
                    var lines = File.ReadAllLines(SettingsFile);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('=');
                        if (parts.Length == 2)
                        {
                            var key = parts[0].Trim();
                            var value = parts[1].Trim();
                            
                            switch (key)
                            {
                                case "LastTemplatePath":
                                    settings.LastTemplatePath = value;
                                    break;
                                case "LastSourceDirectory":
                                    settings.LastSourceDirectory = value;
                                    break;
                                case "LastExportDirectory":
                                    settings.LastExportDirectory = value;
                                    break;
                                case "FlatImport":
                                    if (bool.TryParse(value, out bool flatImport))
                                        settings.FlatImport = flatImport;
                                    break;
                                case "FlatExport":
                                    if (bool.TryParse(value, out bool flatExport))
                                        settings.FlatExport = flatExport;
                                    break;
                                case "LogToTemplateDirectory":
                                    if (bool.TryParse(value, out bool logToTemplate))
                                        settings.LogToTemplateDirectory = logToTemplate;
                                    break;
                                case "EnableDetailedLogging":
                                    if (bool.TryParse(value, out bool enableDetailed))
                                        settings.EnableDetailedLogging = enableDetailed;
                                    break;
                            }
                        }
                    }
                }
            }
            catch
            {
                // If settings can't be loaded, return defaults
            }
            
            return settings;
        }

        public void Save()
        {
            try
            {
                Directory.CreateDirectory(SettingsDirectory);
                
                var lines = new[]
                {
                    $"LastTemplatePath={LastTemplatePath}",
                    $"LastSourceDirectory={LastSourceDirectory}",
                    $"LastExportDirectory={LastExportDirectory}",
                    $"FlatImport={FlatImport}",
                    $"FlatExport={FlatExport}",
                    $"LogToTemplateDirectory={LogToTemplateDirectory}",
                    $"EnableDetailedLogging={EnableDetailedLogging}"
                };
                
                File.WriteAllLines(SettingsFile, lines);
            }
            catch
            {
                // Silently fail if settings can't be saved
            }
        }
    }
}