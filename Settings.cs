using System;
using System.IO;

namespace BuildingBlocksManager
{
    public class ProfilePaths
    {
        public string TemplatePath { get; set; } = "";
        public string SourceDirectory { get; set; } = "";
        public string ExportDirectory { get; set; } = "";
        public string LogDirectory { get; set; } = "";
    }

    public enum ProfileType
    {
        Kestrel,
        Compliance
    }

    public class Settings
    {
        private static readonly string SettingsDirectory =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager");
        private static readonly string SettingsFile = Path.Combine(SettingsDirectory, "settings.txt");

        // Current active profile
        public ProfileType ActiveProfile { get; set; } = ProfileType.Kestrel;

        // Profile-specific paths
        public ProfilePaths KestrelPaths { get; set; } = new ProfilePaths();
        public ProfilePaths CompliancePaths { get; set; } = new ProfilePaths();

        // Shared settings (not profile-specific)
        public string LedgerDirectory { get; set; } = "";
        public bool FlatImport { get; set; } = false;
        public bool FlatExport { get; set; } = false;
        public bool EnableDetailedLogging { get; set; } = true;

        // Helper to get current profile's paths
        public ProfilePaths CurrentPaths => ActiveProfile == ProfileType.Kestrel ? KestrelPaths : CompliancePaths;

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
                        var separatorIndex = line.IndexOf('=');
                        if (separatorIndex > 0)
                        {
                            var key = line.Substring(0, separatorIndex).Trim();
                            var value = separatorIndex < line.Length - 1 ? line.Substring(separatorIndex + 1).Trim() : "";

                            switch (key)
                            {
                                // Active profile
                                case "ActiveProfile":
                                    if (Enum.TryParse<ProfileType>(value, out var profile))
                                        settings.ActiveProfile = profile;
                                    break;

                                // Kestrel profile paths
                                case "Kestrel_TemplatePath":
                                    settings.KestrelPaths.TemplatePath = value;
                                    break;
                                case "Kestrel_SourceDirectory":
                                    settings.KestrelPaths.SourceDirectory = value;
                                    break;
                                case "Kestrel_ExportDirectory":
                                    settings.KestrelPaths.ExportDirectory = value;
                                    break;
                                case "Kestrel_LogDirectory":
                                    settings.KestrelPaths.LogDirectory = value;
                                    break;

                                // Compliance profile paths
                                case "Compliance_TemplatePath":
                                    settings.CompliancePaths.TemplatePath = value;
                                    break;
                                case "Compliance_SourceDirectory":
                                    settings.CompliancePaths.SourceDirectory = value;
                                    break;
                                case "Compliance_ExportDirectory":
                                    settings.CompliancePaths.ExportDirectory = value;
                                    break;
                                case "Compliance_LogDirectory":
                                    settings.CompliancePaths.LogDirectory = value;
                                    break;

                                // Shared settings
                                case "LedgerDirectory":
                                    settings.LedgerDirectory = value;
                                    break;
                                case "FlatImport":
                                    if (bool.TryParse(value, out bool flatImport))
                                        settings.FlatImport = flatImport;
                                    break;
                                case "FlatExport":
                                    if (bool.TryParse(value, out bool flatExport))
                                        settings.FlatExport = flatExport;
                                    break;
                                case "EnableDetailedLogging":
                                    if (bool.TryParse(value, out bool enableDetailed))
                                        settings.EnableDetailedLogging = enableDetailed;
                                    break;

                                // Legacy migration - map old settings to Kestrel profile
                                case "LastTemplatePath":
                                    if (string.IsNullOrEmpty(settings.KestrelPaths.TemplatePath))
                                        settings.KestrelPaths.TemplatePath = value;
                                    break;
                                case "LastSourceDirectory":
                                    if (string.IsNullOrEmpty(settings.KestrelPaths.SourceDirectory))
                                        settings.KestrelPaths.SourceDirectory = value;
                                    break;
                                case "LastExportDirectory":
                                    if (string.IsNullOrEmpty(settings.KestrelPaths.ExportDirectory))
                                        settings.KestrelPaths.ExportDirectory = value;
                                    break;
                                case "LogDirectory":
                                    if (string.IsNullOrEmpty(settings.KestrelPaths.LogDirectory))
                                        settings.KestrelPaths.LogDirectory = value;
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
                    $"ActiveProfile={ActiveProfile}",

                    // Kestrel profile
                    $"Kestrel_TemplatePath={KestrelPaths.TemplatePath}",
                    $"Kestrel_SourceDirectory={KestrelPaths.SourceDirectory}",
                    $"Kestrel_ExportDirectory={KestrelPaths.ExportDirectory}",
                    $"Kestrel_LogDirectory={KestrelPaths.LogDirectory}",

                    // Compliance profile
                    $"Compliance_TemplatePath={CompliancePaths.TemplatePath}",
                    $"Compliance_SourceDirectory={CompliancePaths.SourceDirectory}",
                    $"Compliance_ExportDirectory={CompliancePaths.ExportDirectory}",
                    $"Compliance_LogDirectory={CompliancePaths.LogDirectory}",

                    // Shared settings
                    $"LedgerDirectory={LedgerDirectory}",
                    $"FlatImport={FlatImport}",
                    $"FlatExport={FlatExport}",
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
