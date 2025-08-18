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
        private Dictionary<string, LedgerEntry> removedEntries;

        public BuildingBlockLedger()
        {
            System.Diagnostics.Debug.WriteLine("*** LEDGER CONSTRUCTOR: Creating BuildingBlockLedger ***");
            
            // Always use the top-level BBM_Logs directory (not session subdirectory)
            ledgerDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "BBM_Logs");
            ledgerFile = Path.Combine(ledgerDirectory, "building_blocks_ledger.txt");
            
            System.Diagnostics.Debug.WriteLine($"*** LEDGER CONSTRUCTOR: Will load from {ledgerFile} ***");
            
            ledgerEntries = new Dictionary<string, LedgerEntry>();
            removedEntries = new Dictionary<string, LedgerEntry>();
            Load();
        }

        public BuildingBlockLedger(string logDirectory)
        {
            System.Diagnostics.Debug.WriteLine($"*** LEDGER CONSTRUCTOR(logDir): Creating BuildingBlockLedger with logDirectory: '{logDirectory}' ***");
            
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
            
            System.Diagnostics.Debug.WriteLine($"*** LEDGER CONSTRUCTOR(logDir): Will load from {ledgerFile} ***");
            
            ledgerEntries = new Dictionary<string, LedgerEntry>();
            removedEntries = new Dictionary<string, LedgerEntry>();
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
            var found = ledgerEntries.ContainsKey(key);
            var result = found ? ledgerEntries[key].LastModified : DateTime.MinValue;
            
            // DEBUG: Log ledger lookup details
            System.Diagnostics.Debug.WriteLine($"[LEDGER] Lookup Key: '{key}'");
            System.Diagnostics.Debug.WriteLine($"[LEDGER] Found in ledger: {found}");
            if (found)
            {
                System.Diagnostics.Debug.WriteLine($"[LEDGER] Stored LastModified: {ledgerEntries[key].LastModified}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[LEDGER] Available keys in ledger:");
                foreach (var availableKey in ledgerEntries.Keys.Take(5))
                {
                    System.Diagnostics.Debug.WriteLine($"[LEDGER]   '{availableKey}' -> {ledgerEntries[availableKey].LastModified}");
                }
                if (ledgerEntries.Count > 5)
                {
                    System.Diagnostics.Debug.WriteLine($"[LEDGER]   ... and {ledgerEntries.Count - 5} more entries");
                }
            }
            
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
            
            System.Diagnostics.Debug.WriteLine($"[LEDGER UPDATE] Adding entry: '{key}' -> {lastModified}");
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
                System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Starting to load ledger from: {ledgerFile}");
                
                if (File.Exists(ledgerFile))
                {
                    var lines = File.ReadAllLines(ledgerFile);
                    var currentSection = "active"; // Default to active section
                    
                    System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Found ledger file with {lines.Length} lines");
                    
                    foreach (var line in lines)
                    {
                        System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Processing line: '{line}'");
                        
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                        {
                            System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Skipping comment/empty line");
                            continue; // Skip empty lines and comments
                        }

                        // Check for section headers
                        if (line.Trim().ToLower().Contains("removed building blocks"))
                        {
                            currentSection = "removed";
                            System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Switched to removed section");
                            continue;
                        }
                        else if (line.Trim().ToLower().Contains("active building blocks") || 
                                line.Trim().StartsWith("Name") || 
                                currentSection == "active")
                        {
                            // Stay in or switch to active section
                            if (line.Trim().StartsWith("Name"))
                            {
                                System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Skipping header line");
                                continue; // Skip header line
                            }
                        }

                        var targetCollection = currentSection == "removed" ? removedEntries : ledgerEntries;
                        System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Parsing data line in section: {currentSection}");
                        
                        // Parse space-aligned format: "Name                    Category                2025-08-18 14:26"
                        var dateMatch = System.Text.RegularExpressions.Regex.Match(line, @"\s+(\d{4}-\d{2}-\d{2} \d{2}:\d{2})$");
                        System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Date regex match: {dateMatch.Success}");
                        if (dateMatch.Success)
                        {
                            System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Found date: '{dateMatch.Groups[1].Value}'");
                        }
                        
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
                                
                                // DEBUG: Log successful parsing
                                System.Diagnostics.Debug.WriteLine($"[LEDGER PARSE] Loaded entry: Name='{entry.Name}', Category='{entry.Category}', Key='{key}'");
                            }
                            else
                            {
                                // DEBUG: Log parsing failures
                                System.Diagnostics.Debug.WriteLine($"[LEDGER PARSE] Failed to parse line: '{line}'");
                                System.Diagnostics.Debug.WriteLine($"[LEDGER PARSE] NameAndCategory: '{nameAndCategory}'");
                            }
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Finished loading. Active entries: {ledgerEntries.Count}, Removed entries: {removedEntries.Count}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Ledger file does not exist at: {ledgerFile}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LEDGER LOAD] Exception during load: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"[LEDGER SAVE] Attempting to save ledger to: {ledgerFile}");
                System.Diagnostics.Debug.WriteLine($"[LEDGER SAVE] Directory: {ledgerDirectory}");
                System.Diagnostics.Debug.WriteLine($"[LEDGER SAVE] Entries to save: {ledgerEntries.Count} active, {removedEntries.Count} removed");
                
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
                
                System.Diagnostics.Debug.WriteLine($"[LEDGER SAVE] File written successfully. Lines written: {lines.Count}");
                
                // Debug: Verify file was actually created
                if (File.Exists(ledgerFile))
                {
                    var fileInfo = new FileInfo(ledgerFile);
                    System.Diagnostics.Debug.WriteLine($"[LEDGER SAVE] File confirmed exists. Size: {fileInfo.Length} bytes");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[LEDGER SAVE] ERROR: Ledger file was not created at: {ledgerFile}");
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
        /// Force reload the ledger and return diagnostic information
        /// </summary>
        public string GetDiagnosticInfo()
        {
            var info = new System.Text.StringBuilder();
            info.AppendLine($"Ledger file path: {ledgerFile}");
            info.AppendLine($"File exists: {File.Exists(ledgerFile)}");
            
            if (File.Exists(ledgerFile))
            {
                var lines = File.ReadAllLines(ledgerFile);
                info.AppendLine($"File line count: {lines.Length}");
                info.AppendLine("First 10 lines:");
                for (int i = 0; i < Math.Min(10, lines.Length); i++)
                {
                    info.AppendLine($"  [{i}]: '{lines[i]}'");
                }
            }
            
            info.AppendLine($"Active entries in memory: {ledgerEntries.Count}");
            info.AppendLine($"Removed entries in memory: {removedEntries.Count}");
            
            if (ledgerEntries.Count > 0)
            {
                info.AppendLine("Sample keys in memory:");
                foreach (var key in ledgerEntries.Keys.Take(5))
                {
                    info.AppendLine($"  '{key}'");
                }
            }
            
            return info.ToString();
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