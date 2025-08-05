using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BuildingBlocksManager.Core
{
    public class ImportExportManager
    {
        private readonly FileManager _fileManager;
        private readonly WordManager _wordManager;
        private readonly ImportTracker _importTracker;
        private readonly Logger _logger;

        public class ImportOptions
        {
            public bool FlatImport { get; set; }
            public string FlatImportCategory { get; set; }
            public bool ImportOnlyChanged { get; set; } = true;
            public bool ShowWarningsForNewFiles { get; set; } = true;
        }

        public class ExportOptions
        {
            public bool FlatExport { get; set; }
            public string FlatExportDirectory { get; set; }
            public bool HierarchicalExport { get; set; } = true;
            public string ExportRootDirectory { get; set; }
        }

        public class ImportResult
        {
            public bool Success { get; set; }
            public int ImportedCount { get; set; }
            public int FailedCount { get; set; }
            public int SkippedCount { get; set; }
            public List<string> SuccessfulImports { get; set; } = new List<string>();
            public List<string> FailedImports { get; set; } = new List<string>();
            public List<string> NewFiles { get; set; } = new List<string>();
            public List<string> SkippedFiles { get; set; } = new List<string>();
            public string ErrorMessage { get; set; }
            public TimeSpan ProcessingTime { get; set; }
        }

        public class ExportResult
        {
            public bool Success { get; set; }
            public int ExportedCount { get; set; }
            public int FailedCount { get; set; }
            public List<string> SuccessfulExports { get; set; } = new List<string>();
            public List<string> FailedExports { get; set; } = new List<string>();
            public string ErrorMessage { get; set; }
            public string ExportDirectory { get; set; }
            public TimeSpan ProcessingTime { get; set; }
        }

        public class QueryResult
        {
            public int NewFilesCount { get; set; }
            public int ModifiedFilesCount { get; set; }
            public int UpToDateFilesCount { get; set; }
            public int IgnoredFilesCount { get; set; }
            public int MissingFilesCount { get; set; }
            public List<ImportTracker.FileChangeInfo> DetailedChanges { get; set; } = new List<ImportTracker.FileChangeInfo>();
            public List<string> NewFiles { get; set; } = new List<string>();
            public List<string> MissingFiles { get; set; } = new List<string>();
            public TimeSpan ScanTime { get; set; }
        }

        public event EventHandler<string> ProgressUpdate;
        public event EventHandler<int> ProgressPercentageUpdate;

        public ImportExportManager(FileManager fileManager, WordManager wordManager, ImportTracker importTracker, Logger logger)
        {
            _fileManager = fileManager ?? throw new ArgumentNullException(nameof(fileManager));
            _wordManager = wordManager ?? throw new ArgumentNullException(nameof(wordManager));
            _importTracker = importTracker ?? throw new ArgumentNullException(nameof(importTracker));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public QueryResult QueryDirectory(string sourceDirectory)
        {
            var startTime = DateTime.Now;
            var result = new QueryResult();

            try
            {
                OnProgressUpdate("Scanning directory...");
                
                var scanResult = _fileManager.ScanDirectory(sourceDirectory);
                var changeAnalysis = _importTracker.AnalyzeChanges(scanResult.ValidFiles);
                
                result.NewFilesCount = changeAnalysis.NewFiles.Count;
                result.ModifiedFilesCount = changeAnalysis.ModifiedFiles.Count;
                result.UpToDateFilesCount = changeAnalysis.UpToDateFiles.Count;
                result.IgnoredFilesCount = scanResult.IgnoredFiles.Count;
                result.MissingFilesCount = changeAnalysis.MissingFiles.Count;
                result.NewFiles = changeAnalysis.NewFiles;
                result.MissingFiles = changeAnalysis.MissingFiles;
                result.DetailedChanges = _importTracker.GetDetailedChangeInfo(scanResult.ValidFiles);
                result.ScanTime = DateTime.Now - startTime;

                _logger.LogDirectoryScan(sourceDirectory, scanResult.TotalFilesScanned, 
                    scanResult.ValidFiles.Count, scanResult.InvalidFiles.Count, 
                    scanResult.IgnoredFiles.Count, result.ScanTime);

                if (result.NewFilesCount > 0)
                {
                    var newBuildingBlockNames = scanResult.ValidFiles
                        .Where(f => changeAnalysis.NewFiles.Contains(f.FullPath))
                        .Select(f => f.BuildingBlockName)
                        .ToList();
                    
                    _logger.LogNewFilesDetected(result.NewFiles, newBuildingBlockNames);
                }

                if (result.MissingFilesCount > 0)
                {
                    _logger.LogMissingFiles(result.MissingFiles);
                }

                OnProgressUpdate("Directory analysis complete");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to query directory");
                throw;
            }

            return result;
        }

        public ImportResult BatchImport(string sourceDirectory, string templatePath, ImportOptions options = null)
        {
            var startTime = DateTime.Now;
            var result = new ImportResult();
            options = options ?? new ImportOptions();
            string backupPath = null;

            try
            {
                OnProgressUpdate("Starting batch import...");
                
                // Scan directory
                var scanResult = _fileManager.ScanDirectory(sourceDirectory);
                var changeAnalysis = _importTracker.AnalyzeChanges(scanResult.ValidFiles);

                // Check for new files and show warning if enabled
                if (options.ShowWarningsForNewFiles && changeAnalysis.NewFiles.Any())
                {
                    var newBuildingBlockNames = scanResult.ValidFiles
                        .Where(f => changeAnalysis.NewFiles.Contains(f.FullPath))
                        .Select(f => f.BuildingBlockName)
                        .ToList();

                    if (!ShowNewFileWarning(changeAnalysis.NewFiles, newBuildingBlockNames))
                    {
                        result.ErrorMessage = "Import cancelled by user";
                        return result;
                    }

                    result.NewFiles = changeAnalysis.NewFiles;
                }

                // Create backup
                OnProgressUpdate("Creating template backup...");
                backupPath = _wordManager.CreateBackup(templatePath);
                _logger.LogBackupCreated(templatePath, backupPath);

                // Open template
                _wordManager.OpenTemplate(templatePath);

                // Determine files to import
                var filesToImport = scanResult.ValidFiles.Where(f =>
                {
                    if (!options.ImportOnlyChanged) return true;
                    return changeAnalysis.NewFiles.Contains(f.FullPath) || 
                           changeAnalysis.ModifiedFiles.Contains(f.FullPath);
                }).ToList();

                OnProgressUpdate($"Importing {filesToImport.Count} files...");

                // Import files
                for (int i = 0; i < filesToImport.Count; i++)
                {
                    var fileInfo = filesToImport[i];
                    OnProgressPercentageUpdate((i * 100) / filesToImport.Count);
                    
                    try
                    {
                        var category = options.FlatImport ? options.FlatImportCategory : fileInfo.Category;
                        
                        var importResult = _wordManager.ImportBuildingBlock(
                            fileInfo.FullPath, 
                            fileInfo.BuildingBlockName, 
                            category
                        );

                        var exchangeInfo = new Logger.FileExchangeInfo
                        {
                            Operation = "Import",
                            FilePath = fileInfo.FullPath,
                            BuildingBlockName = fileInfo.BuildingBlockName,
                            Category = category,
                            Success = importResult.Success,
                            ErrorMessage = importResult.ErrorMessage,
                            Timestamp = DateTime.Now
                        };

                        _logger.LogFileExchange(exchangeInfo);

                        if (importResult.Success)
                        {
                            result.SuccessfulImports.Add(fileInfo.FullPath);
                            result.ImportedCount++;
                            _importTracker.UpdateImportTime(fileInfo.FullPath);
                        }
                        else
                        {
                            result.FailedImports.Add($"{fileInfo.FullPath}: {importResult.ErrorMessage}");
                            result.FailedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedImports.Add($"{fileInfo.FullPath}: {ex.Message}");
                        result.FailedCount++;
                        _logger.LogError(ex, $"Failed to import {fileInfo.FileName}");
                    }
                }

                // Save template
                OnProgressUpdate("Saving template...");
                _wordManager.SaveTemplate();

                // Update manifest
                _importTracker.UpdateManifest(scanResult.ValidFiles);

                result.Success = result.ImportedCount > 0 || result.FailedCount == 0;
                result.SkippedCount = scanResult.ValidFiles.Count - filesToImport.Count;
                result.ProcessingTime = DateTime.Now - startTime;

                _logger.LogImportSummary(result.ImportedCount, result.FailedCount, result.SkippedCount, result.ProcessingTime);

                OnProgressUpdate($"Import complete: {result.ImportedCount} imported, {result.FailedCount} failed");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Batch import failed");

                // Attempt rollback if backup exists
                if (!string.IsNullOrEmpty(backupPath))
                {
                    try
                    {
                        _wordManager.RestoreFromBackup(backupPath, templatePath);
                        _logger.LogBackupRestored(backupPath, templatePath);
                        OnProgressUpdate("Template restored from backup due to error");
                    }
                    catch (Exception rollbackEx)
                    {
                        _logger.LogError(rollbackEx, "Failed to restore backup after import failure");
                    }
                }

                OnProgressUpdate($"Import failed: {ex.Message}");
            }

            OnProgressPercentageUpdate(100);
            return result;
        }

        public ImportResult SelectiveImport(string templatePath, ImportOptions options = null)
        {
            var result = new ImportResult();
            options = options ?? new ImportOptions();

            try
            {
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "AutoText Files (AT_*.docx)|AT_*.docx|All Word Documents (*.docx)|*.docx";
                    openFileDialog.Title = "Select AutoText File to Import";
                    openFileDialog.Multiselect = false;

                    if (openFileDialog.ShowDialog() != DialogResult.OK)
                    {
                        result.ErrorMessage = "No file selected";
                        return result;
                    }

                    var filePath = openFileDialog.FileName;
                    var fileName = Path.GetFileName(filePath);

                    // Validate file
                    if (!fileName.StartsWith("AT_", StringComparison.OrdinalIgnoreCase))
                    {
                        result.ErrorMessage = "Selected file must start with 'AT_'";
                        return result;
                    }

                    // Check if file has been modified since last import
                    if (_importTracker.HasBeenImported(filePath) && !_importTracker.HasBeenModifiedSinceImport(filePath))
                    {
                        var lastImport = _importTracker.GetLastImportTime(filePath);
                        var confirmResult = MessageBox.Show(
                            $"File '{fileName}' has not been modified since last import ({lastImport:yyyy-MM-dd HH:mm:ss}).\n\nDo you want to import it anyway?",
                            "File Not Modified",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                        if (confirmResult != DialogResult.Yes)
                        {
                            result.ErrorMessage = "Import cancelled - file not modified";
                            return result;
                        }
                    }

                    return ImportSingleFile(filePath, templatePath, options);
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Selective import failed");
            }

            return result;
        }

        private ImportResult ImportSingleFile(string filePath, string templatePath, ImportOptions options)
        {
            var startTime = DateTime.Now;
            var result = new ImportResult();
            string backupPath = null;

            try
            {
                OnProgressUpdate($"Importing {Path.GetFileName(filePath)}...");

                // Create a temporary FileInfo for this single file
                var tempScanResult = new FileManager.ScanResult();
                tempScanResult.ValidFiles.Add(new FileManager.FileInfo
                {
                    FullPath = filePath,
                    FileName = Path.GetFileName(filePath),
                    BuildingBlockName = Path.GetFileNameWithoutExtension(Path.GetFileName(filePath).Substring(3)).Replace(' ', '_'),
                    Category = options.FlatImport ? options.FlatImportCategory : "InternalAutotext",
                    LastModified = File.GetLastWriteTime(filePath),
                    IsValid = true
                });

                // Create backup
                backupPath = _wordManager.CreateBackup(templatePath);
                _logger.LogBackupCreated(templatePath, backupPath);

                // Open template
                _wordManager.OpenTemplate(templatePath);

                var fileInfo = tempScanResult.ValidFiles[0];
                var importResult = _wordManager.ImportBuildingBlock(
                    fileInfo.FullPath,
                    fileInfo.BuildingBlockName,
                    fileInfo.Category
                );

                var exchangeInfo = new Logger.FileExchangeInfo
                {
                    Operation = "Import",
                    FilePath = fileInfo.FullPath,
                    BuildingBlockName = fileInfo.BuildingBlockName,
                    Category = fileInfo.Category,
                    Success = importResult.Success,
                    ErrorMessage = importResult.ErrorMessage,
                    Timestamp = DateTime.Now
                };

                _logger.LogFileExchange(exchangeInfo);

                if (importResult.Success)
                {
                    _wordManager.SaveTemplate();
                    _importTracker.UpdateImportTime(fileInfo.FullPath);
                    
                    result.Success = true;
                    result.ImportedCount = 1;
                    result.SuccessfulImports.Add(fileInfo.FullPath);
                    
                    OnProgressUpdate("Import successful");
                }
                else
                {
                    result.Success = false;
                    result.FailedCount = 1;
                    result.ErrorMessage = importResult.ErrorMessage;
                    result.FailedImports.Add($"{fileInfo.FullPath}: {importResult.ErrorMessage}");
                    
                    OnProgressUpdate($"Import failed: {importResult.ErrorMessage}");
                }

                result.ProcessingTime = DateTime.Now - startTime;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Single file import failed");

                // Attempt rollback
                if (!string.IsNullOrEmpty(backupPath))
                {
                    try
                    {
                        _wordManager.RestoreFromBackup(backupPath, templatePath);
                        _logger.LogBackupRestored(backupPath, templatePath);
                    }
                    catch (Exception rollbackEx)
                    {
                        _logger.LogError(rollbackEx, "Failed to restore backup after import failure");
                    }
                }
            }

            return result;
        }

        private bool ShowNewFileWarning(List<string> newFiles, List<string> newBuildingBlockNames)
        {
            var message = $"⚠️ WARNING: New AutoText Entries Detected\n\n" +
                         $"The following {newFiles.Count} new files will create new Building Blocks:\n\n";

            for (int i = 0; i < Math.Min(newFiles.Count, newBuildingBlockNames.Count); i++)
            {
                var fileName = Path.GetFileName(newFiles[i]);
                message += $"• {fileName} → Building Block: \"{newBuildingBlockNames[i]}\"\n";
            }

            message += $"\nIMPORTANT: These new autotext entries will require updating\n" +
                      $"the Automator configuration to target them programmatically.\n\n" +
                      $"Please update your Automator scripts to reference:\n";

            foreach (var bbName in newBuildingBlockNames)
            {
                message += $"• \"{bbName}\"\n";
            }

            message += $"\nDo you want to continue with the import?";

            var result = MessageBox.Show(message, "New AutoText Files Detected", 
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);

            return result == DialogResult.Yes;
        }

        public ExportResult BatchExport(string templatePath, ExportOptions options)
        {
            var startTime = DateTime.Now;
            var result = new ExportResult();
            options = options ?? new ExportOptions();

            try
            {
                OnProgressUpdate("Starting batch export...");

                _wordManager.OpenTemplate(templatePath);

                var buildingBlocks = _wordManager.GetBuildingBlocks("InternalAutotext");
                
                if (!buildingBlocks.Any())
                {
                    result.ErrorMessage = "No Building Blocks found with 'InternalAutotext' category";
                    return result;
                }

                OnProgressUpdate($"Exporting {buildingBlocks.Count} Building Blocks...");

                var exportDirectory = options.FlatExport ? options.FlatExportDirectory : options.ExportRootDirectory;
                Directory.CreateDirectory(exportDirectory);
                result.ExportDirectory = exportDirectory;

                for (int i = 0; i < buildingBlocks.Count; i++)
                {
                    var bb = buildingBlocks[i];
                    OnProgressPercentageUpdate((i * 100) / buildingBlocks.Count);

                    try
                    {
                        string outputPath;
                        
                        if (options.FlatExport)
                        {
                            outputPath = Path.Combine(exportDirectory, $"AT_{bb.Name}.docx");
                        }
                        else
                        {
                            // Hierarchical export - recreate directory structure
                            var categoryPath = bb.Category.Replace("InternalAutotext\\", "").Replace("InternalAutotext", "");
                            var folderPath = string.IsNullOrEmpty(categoryPath) ? exportDirectory : 
                                           Path.Combine(exportDirectory, categoryPath.Replace('_', ' '));
                            
                            Directory.CreateDirectory(folderPath);
                            outputPath = Path.Combine(folderPath, $"AT_{bb.Name.Replace('_', ' ')}.docx");
                        }

                        // Handle file name conflicts
                        outputPath = GetUniqueFilePath(outputPath);

                        var exportResult = _wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputPath);

                        var exchangeInfo = new Logger.FileExchangeInfo
                        {
                            Operation = "Export",
                            FilePath = outputPath,
                            BuildingBlockName = bb.Name,
                            Category = bb.Category,
                            Success = exportResult.Success,
                            ErrorMessage = exportResult.ErrorMessage,
                            Timestamp = DateTime.Now
                        };

                        _logger.LogFileExchange(exchangeInfo);

                        if (exportResult.Success)
                        {
                            result.SuccessfulExports.Add(outputPath);
                            result.ExportedCount++;
                        }
                        else
                        {
                            result.FailedExports.Add($"{bb.Name}: {exportResult.ErrorMessage}");
                            result.FailedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedExports.Add($"{bb.Name}: {ex.Message}");
                        result.FailedCount++;
                        _logger.LogError(ex, $"Failed to export Building Block {bb.Name}");
                    }
                }

                result.Success = result.ExportedCount > 0;
                result.ProcessingTime = DateTime.Now - startTime;

                _logger.LogExportSummary(result.ExportedCount, result.FailedCount, result.ProcessingTime, exportDirectory);

                OnProgressUpdate($"Export complete: {result.ExportedCount} exported, {result.FailedCount} failed");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Batch export failed");
                OnProgressUpdate($"Export failed: {ex.Message}");
            }

            OnProgressPercentageUpdate(100);
            return result;
        }

        public ExportResult SelectiveExport(string templatePath, ExportOptions options)
        {
            var result = new ExportResult();
            options = options ?? new ExportOptions();

            try
            {
                _wordManager.OpenTemplate(templatePath);
                var buildingBlocks = _wordManager.GetBuildingBlocks("InternalAutotext");

                if (!buildingBlocks.Any())
                {
                    result.ErrorMessage = "No Building Blocks found with 'InternalAutotext' category";
                    return result;
                }

                // Show selection dialog
                var selectedBlocks = ShowBuildingBlockSelectionDialog(buildingBlocks);
                
                if (!selectedBlocks.Any())
                {
                    result.ErrorMessage = "No Building Blocks selected for export";
                    return result;
                }

                // Create export options for selected items
                var exportOptions = new ExportOptions
                {
                    FlatExport = options.FlatExport,
                    FlatExportDirectory = options.FlatExportDirectory,
                    HierarchicalExport = options.HierarchicalExport,
                    ExportRootDirectory = options.ExportRootDirectory
                };

                return ExportSelectedBuildingBlocks(selectedBlocks, exportOptions);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Selective export failed");
            }

            return result;
        }

        private List<WordManager.BuildingBlockInfo> ShowBuildingBlockSelectionDialog(List<WordManager.BuildingBlockInfo> buildingBlocks)
        {
            var selectedBlocks = new List<WordManager.BuildingBlockInfo>();

            using (var dialog = new Form())
            {
                dialog.Text = "Select Building Blocks to Export";
                dialog.Size = new System.Drawing.Size(600, 500);
                dialog.StartPosition = FormStartPosition.CenterParent;

                var listBox = new CheckedListBox();
                listBox.Dock = DockStyle.Fill;
                listBox.CheckOnClick = true;

                foreach (var bb in buildingBlocks.OrderBy(b => b.Category).ThenBy(b => b.Name))
                {
                    var displayText = $"{bb.Category}\\{bb.Name}";
                    listBox.Items.Add(displayText, false);
                }

                var buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 50;

                var selectAllButton = new Button();
                selectAllButton.Text = "Select All";
                selectAllButton.Location = new System.Drawing.Point(10, 10);
                selectAllButton.Click += (s, e) => {
                    for (int i = 0; i < listBox.Items.Count; i++)
                        listBox.SetItemChecked(i, true);
                };

                var selectNoneButton = new Button();
                selectNoneButton.Text = "Select None";
                selectNoneButton.Location = new System.Drawing.Point(100, 10);
                selectNoneButton.Click += (s, e) => {
                    for (int i = 0; i < listBox.Items.Count; i++)
                        listBox.SetItemChecked(i, false);
                };

                var okButton = new Button();
                okButton.Text = "Export Selected";
                okButton.Location = new System.Drawing.Point(300, 10);
                okButton.DialogResult = DialogResult.OK;

                var cancelButton = new Button();
                cancelButton.Text = "Cancel";
                cancelButton.Location = new System.Drawing.Point(400, 10);
                cancelButton.DialogResult = DialogResult.Cancel;

                buttonPanel.Controls.AddRange(new Control[] { selectAllButton, selectNoneButton, okButton, cancelButton });
                dialog.Controls.AddRange(new Control[] { listBox, buttonPanel });

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    for (int i = 0; i < listBox.CheckedItems.Count; i++)
                    {
                        var checkedIndex = listBox.CheckedIndices[i];
                        selectedBlocks.Add(buildingBlocks[checkedIndex]);
                    }
                }
            }

            return selectedBlocks;
        }

        private ExportResult ExportSelectedBuildingBlocks(List<WordManager.BuildingBlockInfo> selectedBlocks, ExportOptions options)
        {
            var startTime = DateTime.Now;
            var result = new ExportResult();

            try
            {
                OnProgressUpdate($"Exporting {selectedBlocks.Count} selected Building Blocks...");

                var exportDirectory = options.FlatExport ? options.FlatExportDirectory : options.ExportRootDirectory;
                Directory.CreateDirectory(exportDirectory);
                result.ExportDirectory = exportDirectory;

                for (int i = 0; i < selectedBlocks.Count; i++)
                {
                    var bb = selectedBlocks[i];
                    OnProgressPercentageUpdate((i * 100) / selectedBlocks.Count);

                    try
                    {
                        string outputPath;

                        if (options.FlatExport)
                        {
                            outputPath = Path.Combine(exportDirectory, $"AT_{bb.Name}.docx");
                        }
                        else
                        {
                            var categoryPath = bb.Category.Replace("InternalAutotext\\", "").Replace("InternalAutotext", "");
                            var folderPath = string.IsNullOrEmpty(categoryPath) ? exportDirectory :
                                           Path.Combine(exportDirectory, categoryPath.Replace('_', ' '));

                            Directory.CreateDirectory(folderPath);
                            outputPath = Path.Combine(folderPath, $"AT_{bb.Name.Replace('_', ' ')}.docx");
                        }

                        outputPath = GetUniqueFilePath(outputPath);

                        var exportResult = _wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputPath);

                        var exchangeInfo = new Logger.FileExchangeInfo
                        {
                            Operation = "Export",
                            FilePath = outputPath,
                            BuildingBlockName = bb.Name,
                            Category = bb.Category,
                            Success = exportResult.Success,
                            ErrorMessage = exportResult.ErrorMessage,
                            Timestamp = DateTime.Now
                        };

                        _logger.LogFileExchange(exchangeInfo);

                        if (exportResult.Success)
                        {
                            result.SuccessfulExports.Add(outputPath);
                            result.ExportedCount++;
                        }
                        else
                        {
                            result.FailedExports.Add($"{bb.Name}: {exportResult.ErrorMessage}");
                            result.FailedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedExports.Add($"{bb.Name}: {ex.Message}");
                        result.FailedCount++;
                        _logger.LogError(ex, $"Failed to export Building Block {bb.Name}");
                    }
                }

                result.Success = result.ExportedCount > 0;
                result.ProcessingTime = DateTime.Now - startTime;

                _logger.LogExportSummary(result.ExportedCount, result.FailedCount, result.ProcessingTime, exportDirectory);

                OnProgressUpdate($"Selective export complete: {result.ExportedCount} exported, {result.FailedCount} failed");
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Selective export failed");
                OnProgressUpdate($"Export failed: {ex.Message}");
            }

            OnProgressPercentageUpdate(100);
            return result;
        }

        private string GetUniqueFilePath(string basePath)
        {
            if (!File.Exists(basePath))
                return basePath;

            var directory = Path.GetDirectoryName(basePath);
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(basePath);
            var extension = Path.GetExtension(basePath);

            int counter = 2;
            string uniquePath;

            do
            {
                uniquePath = Path.Combine(directory, $"{fileNameWithoutExtension}_{counter}{extension}");
                counter++;
            }
            while (File.Exists(uniquePath));

            return uniquePath;
        }

        private void OnProgressUpdate(string message)
        {
            ProgressUpdate?.Invoke(this, message);
        }

        private void OnProgressPercentageUpdate(int percentage)
        {
            ProgressPercentageUpdate?.Invoke(this, percentage);
        }
    }
}