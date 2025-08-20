using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BuildingBlocksManager
{
    public class BuildingBlockLedger
    {
        public class LedgerEntry
        {
            public string Name { get; set; }
            public string Category { get; set; }
            public DateTime LastModified { get; set; }
            
            public override string ToString()
            {
                return $"{Name} ({Category})";
            }
        }

        public class ChangeAnalysis
        {
            public List<FileManager.FileInfo> NewFiles { get; set; } = new List<FileManager.FileInfo>();
            public List<FileManager.FileInfo> ModifiedFiles { get; set; } = new List<FileManager.FileInfo>();
            public List<FileManager.FileInfo> UnchangedFiles { get; set; } = new List<FileManager.FileInfo>();
            public List<LedgerEntry> RemovedEntries { get; set; } = new List<LedgerEntry>();
            
            public int TotalChangedFiles => NewFiles.Count + ModifiedFiles.Count;
            public int TotalFiles => NewFiles.Count + ModifiedFiles.Count + UnchangedFiles.Count;
            
            public string GetSummary()
            {
                if (TotalChangedFiles == 0)
                {
                    return $"All {TotalFiles} files are up-to-date";
                }
                
                var summary = $"{TotalChangedFiles} of {TotalFiles} files need importing";
                if (NewFiles.Count > 0) summary += $" • {NewFiles.Count} new";
                if (ModifiedFiles.Count > 0) summary += $" • {ModifiedFiles.Count} modified";
                
                return summary;
            }
        }

        private string ledgerDirectory;
        private string ledgerFile;

        private Dictionary<string, LedgerEntry> ledgerEntries;
        private Dictionary<string, LedgerEntry> removedEntries;

        public BuildingBlockLedger()
        {
            // Use configured ledger directory or default
            ledgerDirectory = GetConfiguredLedgerDirectory();
            ledgerFile = Path.Combine(ledgerDirectory, "building_blocks_ledger.txt");
            
            ledgerEntries = new Dictionary<string, LedgerEntry>();
            removedEntries = new Dictionary<string, LedgerEntry>();
            Load();
        }


        /// <summary>
        /// Get the configured ledger directory from settings - NO FALLBACKS
        /// </summary>
        private string GetConfiguredLedgerDirectory()
        {
            try
            {
                var settings = Settings.Load();
                if (!string.IsNullOrEmpty(settings.LedgerDirectory))
                {
                    return settings.LedgerDirectory;
                }
            }
            catch
            {
                // If settings can't be loaded, use default
            }
            
            // Default location - create if it doesn't exist
            var defaultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "BBM_Logs");
            Directory.CreateDirectory(defaultPath);
            return defaultPath;
        }

        /// <summary>
        /// Generate a unique key for a Building Block based on name and category
        /// </summary>
        private string GetLedgerKey(string name, string category)
        {
            // Use pipe separator to create unique key
            return $"{name}|{category ?? ""}";
        }

        /// <summary>
        /// Get the last import time for a specific Building Block
        /// </summary>
        public DateTime GetLastImportTime(string name, string category)
        {
            var key = GetLedgerKey(name, category);
            var found = ledgerEntries.ContainsKey(key);
            var result = found ? ledgerEntries[key].LastModified : DateTime.MinValue;
            
            return result;
        }

        /// <summary>
        /// Update or add a Building Block entry in the ledger
        /// </summary>
        public void UpdateEntry(string name, string category, DateTime lastModified)
        {
            var key = GetLedgerKey(name, category);
            var entry = new LedgerEntry
            {
                Name = name,
                Category = category ?? "",
                LastModified = lastModified
            };
            
            ledgerEntries[key] = entry;
            Save();
        }

        /// <summary>
        /// Update entry with current timestamp
        /// </summary>
        public void UpdateEntry(string name, string category)
        {
            UpdateEntry(name, category, DateTime.Now);
        }

        /// <summary>
        /// Remove a Building Block entry from the active ledger and move it to removed section
        /// </summary>
        public void RemoveEntry(string name, string category)
        {
            var key = GetLedgerKey(name, category);
            if (ledgerEntries.ContainsKey(key))
            {
                // Move entry from active to removed section
                var entry = ledgerEntries[key];
                entry.LastModified = DateTime.Now; // Update timestamp when removed
                removedEntries[key] = entry;
                ledgerEntries.Remove(key);
                Save();
            }
            else
            {
                // Item doesn't exist in active ledger - create new removed entry
                AddRemovedEntry(name, category, DateTime.Now);
            }
        }

        /// <summary>
        /// Add a Building Block directly to the removed section (for items deleted before being tracked)
        /// </summary>
        public void AddRemovedEntry(string name, string category, DateTime removedTime)
        {
            var key = GetLedgerKey(name, category);
            var entry = new LedgerEntry
            {
                Name = name,
                Category = category ?? "",
                LastModified = removedTime
            };
            
            removedEntries[key] = entry;
            Save();
        }

        /// <summary>
        /// Restore a removed Building Block entry back to the active ledger
        /// </summary>
        public void RestoreEntry(string name, string category, DateTime lastModified)
        {
            var key = GetLedgerKey(name, category);
            if (removedEntries.ContainsKey(key))
            {
                // Move entry from removed back to active section
                var entry = removedEntries[key];
                entry.LastModified = lastModified;
                ledgerEntries[key] = entry;
                removedEntries.Remove(key);
                Save();
            }
        }

        /// <summary>
        /// Analyze changes between current directory scan and ledger entries
        /// </summary>
        public ChangeAnalysis AnalyzeChanges(List<FileManager.FileInfo> scannedFiles)
        {
            var analysis = new ChangeAnalysis();
            var foundInSource = new HashSet<string>();

            foreach (var file in scannedFiles)
            {
                var key = GetLedgerKey(file.Name, file.Category);
                foundInSource.Add(key);

                if (!ledgerEntries.ContainsKey(key))
                {
                    // Check if this file was previously removed and should be restored
                    if (removedEntries.ContainsKey(key))
                    {
                        // File was removed but now exists again - restore it and treat as modified
                        RestoreEntry(file.Name, file.Category, file.LastModified);
                        analysis.ModifiedFiles.Add(file);
                    }
                    else
                    {
                        // New Building Block - not in ledger
                        analysis.NewFiles.Add(file);
                    }
                }
                else
                {
                    var ledgerEntry = ledgerEntries[key];
                    
                    // Compare file modification time with ledger entry using 1-minute tolerance
                    var timeDifference = file.LastModified - ledgerEntry.LastModified;
                    if (timeDifference.TotalMinutes > 1.0)
                    {
                        // File has been modified more than 1 minute after last import
                        analysis.ModifiedFiles.Add(file);
                    }
                    else
                    {
                        // File is unchanged (within 1-minute tolerance)
                        analysis.UnchangedFiles.Add(file);
                    }
                }
            }

            // Find entries in ledger that are no longer in source directory
            foreach (var ledgerEntry in ledgerEntries.Values)
            {
                var key = GetLedgerKey(ledgerEntry.Name, ledgerEntry.Category);
                if (!foundInSource.Contains(key))
                {
                    analysis.RemovedEntries.Add(ledgerEntry);
                }
            }

            return analysis;
        }

        /// <summary>
        /// Get all active entries in the ledger
        /// </summary>
        public List<LedgerEntry> GetAllEntries()
        {
            return ledgerEntries.Values.OrderBy(e => e.Name).ToList();
        }

        /// <summary>
        /// Get all removed entries in the ledger
        /// </summary>
        public List<LedgerEntry> GetRemovedEntries()
        {
            return removedEntries.Values.OrderBy(e => e.Name).ToList();
        }

        /// <summary>
        /// Check if a Building Block exists in the ledger
        /// </summary>
        public bool HasEntry(string name, string category)
        {
            var key = GetLedgerKey(name, category);
            return ledgerEntries.ContainsKey(key);
        }

        /// <summary>
        /// Clear all entries from the ledger (both active and removed)
        /// </summary>
        public void Clear()
        {
            ledgerEntries.Clear();
            removedEntries.Clear();
            Save();
        }

        /// <summary>
        /// Load ledger entries from file
        /// </summary>
        private void Load()
        {
            try
            {
                if (File.Exists(ledgerFile))
                {
                    var lines = File.ReadAllLines(ledgerFile);
                    var currentSection = "active"; // Default to active section
                    
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                            continue; // Skip empty lines and comments

                        // Check for section headers
                        if (line.Trim().ToLower().Contains("removed building blocks"))
                        {
                            currentSection = "removed";
                            continue;
                        }
                        else if (line.Trim().ToLower().Contains("active building blocks") || 
                                line.Trim().StartsWith("Name") || 
                                currentSection == "active")
                        {
                            // Stay in or switch to active section
                            if (line.Trim().StartsWith("Name"))
                                continue; // Skip header line
                        }

                        var targetCollection = currentSection == "removed" ? removedEntries : ledgerEntries;
                        
                        // Parse space-aligned format: "Name                    Category                2025-08-18 14:26"
                        var dateMatch = System.Text.RegularExpressions.Regex.Match(line, @"\s+(\d{4}-\d{2}-\d{2} \d{2}:\d{2})$");
                        
                        if (dateMatch.Success && DateTime.TryParse(dateMatch.Groups[1].Value, out DateTime lastModified))
                        {
                            // Remove the date part and its preceding whitespace
                            var nameAndCategory = line.Substring(0, dateMatch.Index).Trim();
                            
                            // Split into words and find the boundary between name and category
                            // The name ends where we find 2+ consecutive spaces (indicating column separation)
                            var nameMatch = System.Text.RegularExpressions.Regex.Match(nameAndCategory, @"^(\S+(?:\s+\S+)*?)\s{2,}(.+)$");
                            if (nameMatch.Success)
                            {
                                var entry = new LedgerEntry
                                {
                                    Name = nameMatch.Groups[1].Value.Trim(),
                                    Category = nameMatch.Groups[2].Value.Trim(),
                                    LastModified = lastModified
                                };
                                
                                var key = GetLedgerKey(entry.Name, entry.Category);
                                targetCollection[key] = entry;
                            }
                        }
                    }
                    
                }
            }
            catch
            {
                // Silently handle errors - start with empty ledger
                ledgerEntries.Clear();
                removedEntries.Clear();
            }
        }

        /// <summary>
        /// Save ledger entries to file
        /// </summary>
        private void Save()
        {
            try
            {
                Directory.CreateDirectory(ledgerDirectory);
                
                var lines = new List<string>
                {
                    "# Building Blocks Ledger",
                    $"# Last Updated: {DateTime.Now:yyyy-MM-dd HH:mm}",
                    $"# Ledger Path: {ledgerFile}",
                    "#",
                    "# Active Building Blocks",
                    "# Name".PadRight(45) + "Category".PadRight(30) + "LastModified"
                };
                
                foreach (var entry in ledgerEntries.Values.OrderBy(e => e.Name))
                {
                    lines.Add($"{entry.Name.PadRight(45)}{entry.Category.PadRight(30)}{entry.LastModified:yyyy-MM-dd HH:mm}");
                }
                
                // Add removed entries section if any exist
                if (removedEntries.Count > 0)
                {
                    lines.Add("");
                    lines.Add("# Removed Building Blocks");
                    lines.Add("# Name".PadRight(45) + "Category".PadRight(30) + "RemovedDate");
                    
                    foreach (var entry in removedEntries.Values.OrderBy(e => e.Name))
                    {
                        lines.Add($"{entry.Name.PadRight(45)}{entry.Category.PadRight(30)}{entry.LastModified:yyyy-MM-dd HH:mm}");
                    }
                }
                
                File.WriteAllLines(ledgerFile, lines);
            }
            catch (Exception ex)
            {
                // Show save errors for debugging
                System.Windows.Forms.MessageBox.Show($"Failed to save Building Block ledger:\n{ex.Message}\n\nPath: {ledgerFile}", "Ledger Save Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// Get the path to the ledger file
        /// </summary>
        public string GetLedgerFilePath()
        {
            return ledgerFile;
        }

        /// <summary>
        /// Get the ledger directory path
        /// </summary>
        public string GetLedgerDirectory()
        {
            return ledgerDirectory;
        }
        
        /// <summary>
        /// Check if the ledger file exists
        /// </summary>
        public bool LedgerFileExists()
        {
            return File.Exists(ledgerFile);
        }

        /// <summary>
        /// Get simple ledger information
        /// </summary>
        public string GetLedgerInfo()
        {
            return $"Directory: {ledgerDirectory}\nFile: {(File.Exists(ledgerFile) ? "Exists" : "Not found")}\nActive entries: {ledgerEntries.Count}\nRemoved entries: {removedEntries.Count}";
        }

        /// <summary>
        /// Generate ledger from current template Building Blocks (user-relevant only)
        /// </summary>
        public void GenerateFromTemplate(string templatePath)
        {
            try
            {
                // Clear existing entries (both active and removed)
                ledgerEntries.Clear();
                removedEntries.Clear();

                // Use WordManager to get Building Blocks from template
                using (var wordManager = new WordManager(templatePath))
                {
                    var buildingBlocks = wordManager.GetBuildingBlocks();
                    var templateModTime = File.GetLastWriteTime(templatePath);
                    
                    // Filter to only user-relevant Building Blocks (same as Template tab)
                    var userBlocks = buildingBlocks.Where(bb => !IsSystemEntry(bb) && bb.Gallery != "Placeholder").ToList();
                    
                    foreach (var bb in userBlocks)
                    {
                        // Use template modification time as baseline
                        UpdateEntry(bb.Name, bb.Category, templateModTime);
                    }
                }
                
                // Save the newly generated ledger
                Save();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to generate ledger from template: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Detect system/internal Building Blocks that should be excluded from ledger
        /// (Same logic as MainForm.IsSystemEntry)
        /// </summary>
        private bool IsSystemEntry(BuildingBlockInfo bb)
        {
            // More targeted system entry detection
            return bb.Name.Length > 15 && bb.Name.All(c => "0123456789ABCDEF-".Contains(char.ToUpper(c))) ||
                   bb.Name.StartsWith("_") ||
                   (bb.Category != null && bb.Category.Contains("System"));
        }
    }
}