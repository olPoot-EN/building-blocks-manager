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
            public string SourceFilePath { get; set; }
            
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
                var summary = $"Found {TotalFiles} total files";
                if (TotalChangedFiles > 0)
                {
                    summary += $"\n• {NewFiles.Count} new files";
                    summary += $"\n• {ModifiedFiles.Count} modified files";
                    summary += $"\n• {UnchangedFiles.Count} unchanged files";
                }
                else
                {
                    summary += "\n• All files are unchanged since last import";
                }
                
                if (RemovedEntries.Count > 0)
                {
                    summary += $"\n• {RemovedEntries.Count} Building Blocks no longer found in source";
                }
                
                return summary;
            }
        }

        private string ledgerDirectory;
        private string ledgerFile;

        private Dictionary<string, LedgerEntry> ledgerEntries;

        public BuildingBlockLedger()
        {
            // Always use the top-level BBM_Logs directory (not session subdirectory)
            ledgerDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "BBM_Logs");
            ledgerFile = Path.Combine(ledgerDirectory, "building_blocks_ledger.txt");
            
            ledgerEntries = new Dictionary<string, LedgerEntry>();
            Load();
        }

        public BuildingBlockLedger(string logDirectory)
        {
            // Extract the base log directory (remove session folder if present)
            // Logger passes session-specific directory, but ledger should be at top level
            if (logDirectory != null && Path.GetFileName(logDirectory).Contains("-"))
            {
                // This looks like a session directory, get parent
                ledgerDirectory = Path.GetDirectoryName(logDirectory);
            }
            else
            {
                // This is already the base log directory
                ledgerDirectory = logDirectory ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "BBM_Logs");
            }
            
            ledgerFile = Path.Combine(ledgerDirectory, "building_blocks_ledger.txt");
            
            ledgerEntries = new Dictionary<string, LedgerEntry>();
            Load();
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
            return ledgerEntries.ContainsKey(key) ? ledgerEntries[key].LastModified : DateTime.MinValue;
        }

        /// <summary>
        /// Update or add a Building Block entry in the ledger
        /// </summary>
        public void UpdateEntry(string name, string category, DateTime lastModified, string sourceFilePath)
        {
            var key = GetLedgerKey(name, category);
            var entry = new LedgerEntry
            {
                Name = name,
                Category = category ?? "",
                LastModified = lastModified,
                SourceFilePath = sourceFilePath
            };
            
            ledgerEntries[key] = entry;
            Save();
        }

        /// <summary>
        /// Update entry with current timestamp
        /// </summary>
        public void UpdateEntry(string name, string category, string sourceFilePath)
        {
            UpdateEntry(name, category, DateTime.Now, sourceFilePath);
        }

        /// <summary>
        /// Remove a Building Block entry from the ledger
        /// </summary>
        public void RemoveEntry(string name, string category)
        {
            var key = GetLedgerKey(name, category);
            if (ledgerEntries.ContainsKey(key))
            {
                ledgerEntries.Remove(key);
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
                    // New Building Block - not in ledger
                    analysis.NewFiles.Add(file);
                }
                else
                {
                    var ledgerEntry = ledgerEntries[key];
                    
                    // Compare file modification time with ledger entry
                    if (file.LastModified > ledgerEntry.LastModified)
                    {
                        // File has been modified since last import
                        analysis.ModifiedFiles.Add(file);
                    }
                    else
                    {
                        // File is unchanged
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
        /// Get all entries in the ledger
        /// </summary>
        public List<LedgerEntry> GetAllEntries()
        {
            return ledgerEntries.Values.OrderBy(e => e.Name).ToList();
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
        /// Clear all entries from the ledger
        /// </summary>
        public void Clear()
        {
            ledgerEntries.Clear();
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
                    foreach (var line in lines)
                    {
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                            continue; // Skip empty lines and comments

                        var parts = line.Split('|');
                        if (parts.Length >= 4 && DateTime.TryParse(parts[2], out DateTime lastModified))
                        {
                            var entry = new LedgerEntry
                            {
                                Name = parts[0],
                                Category = parts[1],
                                LastModified = lastModified,
                                SourceFilePath = parts[3]
                            };
                            
                            var key = GetLedgerKey(entry.Name, entry.Category);
                            ledgerEntries[key] = entry;
                        }
                    }
                }
            }
            catch
            {
                // Silently handle errors - start with empty ledger
                ledgerEntries.Clear();
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
                    "# Format: Name|Category|LastModified|SourceFilePath",
                    $"# Last Updated: {DateTime.Now:yyyy-MM-dd HH:mm}",
                    $"# Ledger Path: {ledgerFile}",
                    ""
                };
                
                foreach (var entry in ledgerEntries.Values.OrderBy(e => e.Name))
                {
                    lines.Add($"{entry.Name}|{entry.Category}|{entry.LastModified:yyyy-MM-dd HH:mm}|{entry.SourceFilePath}");
                }
                
                File.WriteAllLines(ledgerFile, lines);
                
                // Debug: Verify file was actually created
                if (!File.Exists(ledgerFile))
                {
                    System.Diagnostics.Debug.WriteLine($"ERROR: Ledger file was not created at: {ledgerFile}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"SUCCESS: Ledger saved to: {ledgerFile}");
                }
            }
            catch (Exception ex)
            {
                // Show save errors for debugging
                System.Diagnostics.Debug.WriteLine($"ERROR saving ledger to {ledgerFile}: {ex.Message}");
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
        /// Generate ledger from current template Building Blocks (user-relevant only)
        /// </summary>
        public void GenerateFromTemplate(string templatePath)
        {
            try
            {
                // Clear existing entries
                ledgerEntries.Clear();

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
                        // Source path is the template since these come from template
                        UpdateEntry(bb.Name, bb.Category, templateModTime, templatePath);
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