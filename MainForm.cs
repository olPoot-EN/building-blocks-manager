using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BuildingBlocksManager
{
    public partial class MainForm : Form
    {
        private TextBox txtTemplatePath;
        private TextBox txtSourceDirectory;
        private CheckBox chkFlatImport;
        private CheckBox chkFlatExport;
        private Button btnBrowseTemplate;
        private Button btnBrowseDirectory;
        private Button btnQueryDirectory;
        private Button btnImportAll;
        private Button btnImportSelected;
        private Button btnExportAll;
        private Button btnExportSelected;
        private TextBox txtResults;
        private TabControl tabControl;
        private TabPage tabResults;
        private TabPage tabDirectory;
        private TabPage tabTemplate;
        private TreeView treeDirectory;
        private ListBox listTemplate;
        private ComboBox cmbCategoryFilter;
        private Button btnLoadTemplate;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Settings settings;
        private Logger logger;

        public MainForm()
        {
            InitializeComponent();
            this.Text = "Building Blocks Manager";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(600, 500);
            
            // Wire up event handlers
            btnBrowseTemplate.Click += BtnBrowseTemplate_Click;
            btnBrowseDirectory.Click += BtnBrowseDirectory_Click;
            btnQueryDirectory.Click += BtnQueryDirectory_Click;
            btnImportAll.Click += BtnImportAll_Click;
            btnImportSelected.Click += BtnImportSelected_Click;
            btnExportAll.Click += BtnExportAll_Click;
            btnExportSelected.Click += BtnExportSelected_Click;
            btnLoadTemplate.Click += BtnLoadTemplate_Click;
            cmbCategoryFilter.SelectedIndexChanged += CmbCategoryFilter_SelectedIndexChanged;
            
            // Load and apply settings
            LoadSettings();
            
            // Initialize logger and clean up old logs
            logger = new Logger();
            Logger.CleanupOldLogs();
            logger.Info("Building Blocks Manager started");
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // Create menu bar
            var menuStrip = new MenuStrip();
            
            var fileMenu = new ToolStripMenuItem("File");
            var viewLogMenuItem = new ToolStripMenuItem("View Log File");
            viewLogMenuItem.Click += ViewLogMenuItem_Click;
            fileMenu.DropDownItems.Add(viewLogMenuItem);
            
            fileMenu.DropDownItems.Add(new ToolStripSeparator());
            
            var rollbackMenuItem = new ToolStripMenuItem("Rollback");
            rollbackMenuItem.Click += BtnRollback_Click;
            fileMenu.DropDownItems.Add(rollbackMenuItem);
            
            menuStrip.Items.Add(fileMenu);
            
            this.MainMenuStrip = menuStrip;
            this.Controls.Add(menuStrip);

            // Template file section (adjusted for menu bar)
            var lblTemplate = new Label
            {
                Text = "Template File:",
                Location = new System.Drawing.Point(20, 45),
                Size = new System.Drawing.Size(100, 23)
            };

            txtTemplatePath = new TextBox
            {
                Location = new System.Drawing.Point(130, 45),
                Size = new System.Drawing.Size(500, 23),
                ReadOnly = true
            };

            btnBrowseTemplate = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(640, 44),
                Size = new System.Drawing.Size(80, 25)
            };

            // Source directory section
            var lblDirectory = new Label
            {
                Text = "Source Directory:",
                Location = new System.Drawing.Point(20, 85),
                Size = new System.Drawing.Size(110, 23)
            };

            txtSourceDirectory = new TextBox
            {
                Location = new System.Drawing.Point(140, 85),
                Size = new System.Drawing.Size(490, 23),
                ReadOnly = true
            };

            btnBrowseDirectory = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(640, 84),
                Size = new System.Drawing.Size(80, 25)
            };

            // Structure options section
            var lblStructure = new Label
            {
                Text = "Ignore folder/category structure for:",
                Location = new System.Drawing.Point(20, 125),
                Size = new System.Drawing.Size(250, 23)
            };

            chkFlatImport = new CheckBox
            {
                Text = "Import",
                Location = new System.Drawing.Point(280, 125),
                Size = new System.Drawing.Size(80, 23)
            };

            chkFlatExport = new CheckBox
            {
                Text = "Export",
                Location = new System.Drawing.Point(370, 125),
                Size = new System.Drawing.Size(80, 23)
            };

            // Query group
            var lblQuery = new Label
            {
                Text = "Query",
                Location = new System.Drawing.Point(20, 165),
                Size = new System.Drawing.Size(100, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            btnQueryDirectory = new Button
            {
                Text = "Query Source Directory",
                Location = new System.Drawing.Point(20, 190),
                Size = new System.Drawing.Size(160, 30)
            };

            // Import group
            var lblImport = new Label
            {
                Text = "Import\n(Folder -> Template)",
                Location = new System.Drawing.Point(220, 155),
                Size = new System.Drawing.Size(120, 35),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            btnImportAll = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(220, 195),
                Size = new System.Drawing.Size(80, 30)
            };

            btnImportSelected = new Button
            {
                Text = "Selected",
                Location = new System.Drawing.Point(220, 230),
                Size = new System.Drawing.Size(80, 30)
            };

            // Export group
            var lblExport = new Label
            {
                Text = "Export\n(Template -> Folder)",
                Location = new System.Drawing.Point(350, 155),
                Size = new System.Drawing.Size(120, 35),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            btnExportAll = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(350, 195),
                Size = new System.Drawing.Size(80, 30)
            };

            btnExportSelected = new Button
            {
                Text = "Selected",
                Location = new System.Drawing.Point(350, 230),
                Size = new System.Drawing.Size(80, 30)
            };

            // Tab control section
            tabControl = new TabControl
            {
                Location = new System.Drawing.Point(20, 280),
                Size = new System.Drawing.Size(740, 245),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            // Results tab
            tabResults = new TabPage("Results");
            txtResults = new TextBox
            {
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(725, 210),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            tabResults.Controls.Add(txtResults);

            // Directory tab
            tabDirectory = new TabPage("Directory");
            treeDirectory = new TreeView
            {
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(725, 210),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };
            tabDirectory.Controls.Add(treeDirectory);

            // Template tab
            tabTemplate = new TabPage("Template");
            
            btnLoadTemplate = new Button
            {
                Text = "Load Template",
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(100, 25)
            };

            var lblCategoryFilter = new Label
            {
                Text = "Category:",
                Location = new System.Drawing.Point(115, 10),
                Size = new System.Drawing.Size(60, 15)
            };

            cmbCategoryFilter = new ComboBox
            {
                Location = new System.Drawing.Point(180, 7),
                Size = new System.Drawing.Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };

            listTemplate = new ListBox
            {
                Location = new System.Drawing.Point(5, 35),
                Size = new System.Drawing.Size(725, 175),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom
            };

            tabTemplate.Controls.AddRange(new Control[] { btnLoadTemplate, lblCategoryFilter, cmbCategoryFilter, listTemplate });

            // Add tabs to tab control
            tabControl.TabPages.AddRange(new TabPage[] { tabResults, tabDirectory, tabTemplate });

            // Progress and status section
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 535),
                Size = new System.Drawing.Size(520, 23),
                Style = ProgressBarStyle.Continuous
            };

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(550, 535),
                Size = new System.Drawing.Size(210, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            // Add all controls to form
            Controls.AddRange(new Control[]
            {
                lblTemplate, txtTemplatePath, btnBrowseTemplate,
                lblDirectory, txtSourceDirectory, btnBrowseDirectory,
                lblStructure, chkFlatImport, chkFlatExport,
                lblQuery, btnQueryDirectory, 
                lblImport, btnImportAll, btnImportSelected,
                lblExport, btnExportAll, btnExportSelected,
                tabControl,
                progressBar, lblStatus
            });

            ResumeLayout(false);
        }

        private void BtnBrowseTemplate_Click(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select Word Template File";
                dialog.Filter = "Word Template Files (*.dotm)|*.dotm|All Files (*.*)|*.*";
                dialog.CheckFileExists = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtTemplatePath.Text = dialog.FileName;
                    UpdateStatus("Template file selected: " + Path.GetFileName(dialog.FileName));
                    SaveSettings();
                }
            }
        }

        private void BtnBrowseDirectory_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Source Directory";
                dialog.ShowNewFolderButton = false;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtSourceDirectory.Text = dialog.SelectedPath;
                    UpdateStatus("Source directory selected: " + dialog.SelectedPath);
                    SaveSettings();
                }
            }
        }

        private void UpdateStatus(string message)
        {
            lblStatus.Text = message;
            Application.DoEvents();
        }

        private void AppendResults(string message)
        {
            txtResults.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\r\n");
            txtResults.ScrollToCaret();
            Application.DoEvents();
        }

        private bool ValidatePaths()
        {
            if (string.IsNullOrWhiteSpace(txtTemplatePath.Text))
            {
                MessageBox.Show("Please select a template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!File.Exists(txtTemplatePath.Text))
            {
                MessageBox.Show("The selected template file does not exist.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSourceDirectory.Text))
            {
                MessageBox.Show("Please select a source directory.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!Directory.Exists(txtSourceDirectory.Text))
            {
                MessageBox.Show("The selected source directory does not exist.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        private void BtnQueryDirectory_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            UpdateStatus("Querying directory...");
            progressBar.Style = ProgressBarStyle.Marquee;
            
            // Clear previous directory tree
            treeDirectory.Nodes.Clear();
            
            AppendResults("=== DIRECTORY QUERY ===");
            AppendResults($"Scanning directory: {txtSourceDirectory.Text}");
            
            try
            {
                var fileManager = new FileManager(txtSourceDirectory.Text);
                var summary = fileManager.GetSummary();
                
                AppendResults("");
                AppendResults(summary);
                
                var files = fileManager.ScanDirectory();
                
                // Populate Directory tab tree view
                PopulateDirectoryTree(files);
                
                // Switch to Directory tab
                tabControl.SelectedTab = tabDirectory;
                
                AppendResults("");
                AppendResults($"Directory tree populated with {files.Count} files.");
                AppendResults("Query completed successfully.");
            }
            catch (Exception ex)
            {
                AppendResults($"Error during directory scan: {ex.Message}");
            }
            
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            UpdateStatus("Query completed");
        }

        private void BtnImportAll_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            string flatCategory = null;
            if (chkFlatImport.Checked)
            {
                flatCategory = PromptForInput("Flat Import Category", 
                    "Enter category name for all imported Building Blocks:");
                if (string.IsNullOrWhiteSpace(flatCategory)) return;
                AppendResults($"Using flat import category: InternalAutotext\\{flatCategory}");
            }

            UpdateStatus("Importing all files...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            
            AppendResults("=== IMPORT ALL OPERATION ===");
            AppendResults($"Template: {Path.GetFileName(txtTemplatePath.Text)}");
            AppendResults($"Source Directory: {txtSourceDirectory.Text}");
            AppendResults($"Flat Import: {(chkFlatImport.Checked ? "Yes" : "No")}");
            AppendResults("");
            
            logger.Info($"Starting Import All operation - Template: {txtTemplatePath.Text}, Directory: {txtSourceDirectory.Text}");

            WordManager wordManager = null;
            var startTime = DateTime.Now;
            int successCount = 0;
            int errorCount = 0;

            try
            {
                // Initialize managers
                wordManager = new WordManager(txtTemplatePath.Text);
                var fileManager = new FileManager(txtSourceDirectory.Text);
                var importTracker = new ImportTracker();

                // Create backup
                AppendResults("Creating backup...");
                wordManager.CreateBackup();

                // Get files to import (new and modified only)
                var filesToImport = fileManager.ScanDirectory()
                    .Where(f => (f.IsNew || f.IsModified) && f.IsValid)
                    .ToList();

                if (filesToImport.Count == 0)
                {
                    AppendResults("No files require importing.");
                    return;
                }

                AppendResults($"Found {filesToImport.Count} files to import");
                AppendResults("");

                // Import each file
                for (int i = 0; i < filesToImport.Count; i++)
                {
                    var file = filesToImport[i];
                    var fileName = Path.GetFileName(file.FilePath);
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / filesToImport.Count * 100);
                        UpdateStatus($"Processing {fileName}...");
                        AppendResults($"Processing {fileName}...");

                        // Use flat category if specified, otherwise use extracted category
                        var category = chkFlatImport.Checked ? $"InternalAutotext\\{flatCategory}" : file.Category;
                        
                        // Import the Building Block
                        wordManager.ImportBuildingBlock(file.FilePath, category, file.Name);
                        
                        // Update import tracking
                        importTracker.UpdateImportTime(file.FilePath);
                        
                        successCount++;
                        AppendResults($"  ✓ Successfully imported as {category}\\{file.Name}");
                        logger.Success($"Imported {fileName} as {category}\\{file.Name}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        AppendResults($"  ✗ Failed to import {fileName}: {ex.Message}");
                        logger.Error($"Failed to import {fileName}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Fatal error during import: {ex.Message}");
                UpdateStatus("Import failed");
                return;
            }
            finally
            {
                wordManager?.Dispose();
            }

            var processingTime = (DateTime.Now - startTime).TotalSeconds;
            
            AppendResults("");
            AppendResults("Import Operation Completed");
            AppendResults($"Building Blocks Successfully Imported: {successCount}");
            AppendResults($"Files with Errors: {errorCount}");
            AppendResults($"Processing Time: {processingTime:F1} seconds");
            
            logger.Info($"Import All completed - Success: {successCount}, Errors: {errorCount}, Time: {processingTime:F1}s");
            
            progressBar.Value = 0;
            UpdateStatus(errorCount == 0 ? "Import completed successfully" : $"Import completed with {errorCount} errors");
        }

        private void BtnImportSelected_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Select AT_ File to Import";
                dialog.Filter = "Word Documents (AT_*.docx)|AT_*.docx|All Files (*.*)|*.*";
                dialog.InitialDirectory = txtSourceDirectory.Text;
                dialog.CheckFileExists = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    UpdateStatus("Importing selected file...");
                    progressBar.Style = ProgressBarStyle.Marquee;
                    
                    AppendResults("=== IMPORT SELECTED FILE ===");
                    AppendResults($"File: {Path.GetFileName(dialog.FileName)}");
                    AppendResults($"Template: {Path.GetFileName(txtTemplatePath.Text)}");
                    AppendResults("");

                    WordManager wordManager = null;
                    var startTime = DateTime.Now;
                    
                    try
                    {
                        // Initialize managers
                        wordManager = new WordManager(txtTemplatePath.Text);
                        var fileManager = new FileManager(txtSourceDirectory.Text);
                        var importTracker = new ImportTracker();

                        // Create backup
                        AppendResults("Creating backup...");
                        wordManager.CreateBackup();

                        // Extract category and name
                        var fileName = Path.GetFileName(dialog.FileName);
                        var category = fileManager.ExtractCategory(dialog.FileName);
                        var name = fileManager.ExtractName(fileName);

                        // Check for invalid characters
                        var invalidChars = fileManager.GetInvalidCharacters(fileName);
                        if (invalidChars.Count > 0)
                        {
                            AppendResults($"Warning: File contains invalid characters: {string.Join(", ", invalidChars)}");
                            AppendResults("Import may fail or produce unexpected results.");
                        }

                        AppendResults("Processing file...");
                        
                        // Use flat category if specified
                        if (chkFlatImport.Checked)
                        {
                            string flatCategory = PromptForInput("Flat Import Category", 
                                "Enter category name for this Building Block:");
                            if (string.IsNullOrWhiteSpace(flatCategory))
                            {
                                AppendResults("Import cancelled - no category specified.");
                                return;
                            }
                            category = $"InternalAutotext\\{flatCategory}";
                        }

                        // Import the Building Block
                        wordManager.ImportBuildingBlock(dialog.FileName, category, name);
                        
                        // Update import tracking
                        importTracker.UpdateImportTime(dialog.FileName);
                        
                        var processingTime = (DateTime.Now - startTime).TotalSeconds;
                        
                        AppendResults($"Successfully imported Building Block: {category}\\{name}");
                        AppendResults($"Processing Time: {processingTime:F1} seconds");
                        
                        UpdateStatus("Import completed successfully");
                    }
                    catch (Exception ex)
                    {
                        AppendResults($"Import failed: {ex.Message}");
                        UpdateStatus("Import failed");
                    }
                    finally
                    {
                        wordManager?.Dispose();
                    }
                    
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.Value = 0;
                }
            }
        }

        private void BtnExportAll_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            string exportPath = txtSourceDirectory.Text;
            
            if (chkFlatExport.Checked)
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Select Flat Export Folder";
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    exportPath = dialog.SelectedPath;
                }
            }

            UpdateStatus("Exporting all Building Blocks...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            
            AppendResults("=== EXPORT ALL OPERATION ===");
            AppendResults($"Template: {Path.GetFileName(txtTemplatePath.Text)}");
            AppendResults($"Export Location: {exportPath}");
            AppendResults($"Flat Export: {(chkFlatExport.Checked ? "Yes" : "No")}");
            AppendResults("");

            WordManager wordManager = null;
            var startTime = DateTime.Now;
            int successCount = 0;
            int errorCount = 0;
            int directoriesCreated = 0;

            try
            {
                // Initialize WordManager
                wordManager = new WordManager(txtTemplatePath.Text);
                
                // Get all Building Blocks from template
                AppendResults("Loading Building Blocks from template...");
                var buildingBlocks = wordManager.GetBuildingBlocks();

                if (buildingBlocks.Count == 0)
                {
                    AppendResults("No Building Blocks found in template.");
                    return;
                }

                AppendResults($"Found {buildingBlocks.Count} Building Blocks to export");
                AppendResults("");

                // Export each Building Block
                for (int i = 0; i < buildingBlocks.Count; i++)
                {
                    var bb = buildingBlocks[i];
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / buildingBlocks.Count * 100);
                        UpdateStatus($"Exporting {bb.Name}...");
                        AppendResults($"Exporting {bb.Name}...");

                        string outputFilePath;
                        
                        if (chkFlatExport.Checked)
                        {
                            // Flat export - all files in selected folder
                            outputFilePath = Path.Combine(exportPath, $"AT_{bb.Name}.docx");
                        }
                        else
                        {
                            // Hierarchical export - recreate folder structure
                            var relativePath = ConvertCategoryToPath(bb.Category);
                            var fullDir = Path.Combine(exportPath, relativePath);
                            
                            if (!Directory.Exists(fullDir))
                            {
                                Directory.CreateDirectory(fullDir);
                                directoriesCreated++;
                            }
                            
                            outputFilePath = Path.Combine(fullDir, $"AT_{bb.Name}.docx");
                        }

                        // Handle duplicate filenames
                        outputFilePath = GetUniqueFilePath(outputFilePath);

                        // Export the Building Block
                        wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputFilePath);
                        
                        successCount++;
                        var displayPath = Path.GetRelativePath(exportPath, outputFilePath);
                        AppendResults($"  ✓ Exported to {displayPath}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        AppendResults($"  ✗ Failed to export {bb.Name}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Fatal error during export: {ex.Message}");
                UpdateStatus("Export failed");
                return;
            }
            finally
            {
                wordManager?.Dispose();
            }

            var processingTime = (DateTime.Now - startTime).TotalSeconds;
            
            AppendResults("");
            AppendResults("Export Operation Completed");
            AppendResults($"Building Blocks Successfully Exported: {successCount}");
            AppendResults($"Files with Errors: {errorCount}");
            AppendResults($"Directories Created: {directoriesCreated}");
            AppendResults($"Processing Time: {processingTime:F1} seconds");
            
            progressBar.Value = 0;
            UpdateStatus(errorCount == 0 ? "Export completed successfully" : $"Export completed with {errorCount} errors");
        }

        private void BtnExportSelected_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            // Show Building Block selection dialog
            var selectedBlocks = ShowBuildingBlockSelectionDialog();
            if (selectedBlocks == null || selectedBlocks.Count == 0) return;

            string exportPath = txtSourceDirectory.Text;
            
            if (chkFlatExport.Checked)
            {
                using (var dialog = new FolderBrowserDialog())
                {
                    dialog.Description = "Select Flat Export Folder";
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    exportPath = dialog.SelectedPath;
                }
            }

            UpdateStatus("Exporting selected Building Blocks...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            
            AppendResults("=== EXPORT SELECTED OPERATION ===");
            AppendResults($"Selected Building Blocks: {selectedBlocks.Count}");
            AppendResults($"Export Location: {exportPath}");
            AppendResults("");

            WordManager wordManager = null;
            var startTime = DateTime.Now;
            int successCount = 0;
            int errorCount = 0;
            int directoriesCreated = 0;

            try
            {
                // Initialize WordManager
                wordManager = new WordManager(txtTemplatePath.Text);

                // Export each selected Building Block
                for (int i = 0; i < selectedBlocks.Count; i++)
                {
                    var bb = selectedBlocks[i];
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / selectedBlocks.Count * 100);
                        UpdateStatus($"Exporting {bb.Name}...");
                        AppendResults($"Exporting {bb.Name}...");

                        string outputFilePath;
                        
                        if (chkFlatExport.Checked)
                        {
                            // Flat export - all files in selected folder
                            outputFilePath = Path.Combine(exportPath, $"AT_{bb.Name}.docx");
                        }
                        else
                        {
                            // Hierarchical export - recreate folder structure
                            var relativePath = ConvertCategoryToPath(bb.Category);
                            var fullDir = Path.Combine(exportPath, relativePath);
                            
                            if (!Directory.Exists(fullDir))
                            {
                                Directory.CreateDirectory(fullDir);
                                directoriesCreated++;
                            }
                            
                            outputFilePath = Path.Combine(fullDir, $"AT_{bb.Name}.docx");
                        }

                        // Handle duplicate filenames
                        outputFilePath = GetUniqueFilePath(outputFilePath);

                        // Export the Building Block
                        wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputFilePath);
                        
                        successCount++;
                        var displayPath = Path.GetRelativePath(exportPath, outputFilePath);
                        AppendResults($"  ✓ Exported to {displayPath}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        AppendResults($"  ✗ Failed to export {bb.Name}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Fatal error during export: {ex.Message}");
                UpdateStatus("Export failed");
                return;
            }
            finally
            {
                wordManager?.Dispose();
            }

            var processingTime = (DateTime.Now - startTime).TotalSeconds;
            
            AppendResults("");
            AppendResults("Export Selected Operation Completed");
            AppendResults($"Building Blocks Successfully Exported: {successCount}");
            AppendResults($"Files with Errors: {errorCount}");
            AppendResults($"Directories Created: {directoriesCreated}");
            AppendResults($"Processing Time: {processingTime:F1} seconds");
            
            progressBar.Value = 0;
            UpdateStatus(errorCount == 0 ? "Export completed successfully" : $"Export completed with {errorCount} errors");
        }

        private void BtnRollback_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTemplatePath.Text))
            {
                MessageBox.Show("Please select a template file first.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var result = MessageBox.Show(
                "This will restore the template from the most recent backup. Continue?",
                "Confirm Rollback",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);
                
            if (result == DialogResult.Yes)
            {
                UpdateStatus("Rolling back template...");
                progressBar.Style = ProgressBarStyle.Marquee;
                
                AppendResults("=== ROLLBACK OPERATION ===");
                AppendResults($"Template: {Path.GetFileName(txtTemplatePath.Text)}");
                AppendResults("Searching for backup files...");

                WordManager wordManager = null;
                
                try
                {
                    wordManager = new WordManager(txtTemplatePath.Text);
                    
                    // Perform the rollback
                    wordManager.RollbackFromBackup();
                    
                    AppendResults("Backup file found and restored successfully");
                    AppendResults("Template restored to previous state");
                    AppendResults("Rollback completed successfully");
                    
                    UpdateStatus("Rollback completed");
                }
                catch (Exception ex)
                {
                    AppendResults($"Rollback failed: {ex.Message}");
                    UpdateStatus("Rollback failed");
                }
                finally
                {
                    wordManager?.Dispose();
                }
                
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
            }
        }

        private string PromptForInput(string title, string prompt)
        {
            // Simple input dialog simulation
            return Microsoft.VisualBasic.Interaction.InputBox(prompt, title, "");
        }

        private System.Collections.Generic.List<BuildingBlockInfo> ShowBuildingBlockSelectionDialog()
        {
            // Get real Building Blocks from template
            System.Collections.Generic.List<BuildingBlockInfo> availableBlocks;
            
            try
            {
                using (var wordManager = new WordManager(txtTemplatePath.Text))
                {
                    availableBlocks = wordManager.GetBuildingBlocks();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to load Building Blocks from template: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new System.Collections.Generic.List<BuildingBlockInfo>();
            }

            if (availableBlocks.Count == 0)
            {
                MessageBox.Show("No Building Blocks found in the template.", 
                    "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return new System.Collections.Generic.List<BuildingBlockInfo>();
            }

            // Create selection dialog
            var form = new Form
            {
                Text = "Select Building Blocks to Export",
                Size = new System.Drawing.Size(600, 450),
                StartPosition = FormStartPosition.CenterScreen
            };

            var listBox = new CheckedListBox
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(540, 320),
                CheckOnClick = true
            };

            // Add real Building Blocks to the list
            foreach (var bb in availableBlocks)
            {
                listBox.Items.Add(bb, false);
            }

            var btnSelectAll = new Button
            {
                Text = "Select All",
                Location = new System.Drawing.Point(20, 360),
                Size = new System.Drawing.Size(80, 25)
            };

            var btnSelectNone = new Button
            {
                Text = "Select None",
                Location = new System.Drawing.Point(110, 360),
                Size = new System.Drawing.Size(80, 25)
            };

            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(400, 360),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(490, 360),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.Cancel
            };

            btnSelectAll.Click += (s, e) => {
                for (int i = 0; i < listBox.Items.Count; i++)
                    listBox.SetItemChecked(i, true);
            };

            btnSelectNone.Click += (s, e) => {
                for (int i = 0; i < listBox.Items.Count; i++)
                    listBox.SetItemChecked(i, false);
            };

            form.Controls.AddRange(new Control[] { listBox, btnSelectAll, btnSelectNone, btnOK, btnCancel });
            form.AcceptButton = btnOK;
            form.CancelButton = btnCancel;

            var result = new System.Collections.Generic.List<BuildingBlockInfo>();
            if (form.ShowDialog() == DialogResult.OK)
            {
                foreach (BuildingBlockInfo item in listBox.CheckedItems)
                {
                    result.Add(item);
                }
            }

            form.Dispose();
            return result;
        }

        private void LoadSettings()
        {
            settings = Settings.Load();
            
            txtTemplatePath.Text = settings.LastTemplatePath;
            txtSourceDirectory.Text = settings.LastSourceDirectory;
            chkFlatImport.Checked = settings.FlatImport;
            chkFlatExport.Checked = settings.FlatExport;
            
            if (!string.IsNullOrEmpty(settings.LastTemplatePath) || !string.IsNullOrEmpty(settings.LastSourceDirectory))
            {
                UpdateStatus("Settings loaded - previous paths restored");
            }
        }

        private void SaveSettings()
        {
            if (settings == null) settings = new Settings();
            
            settings.LastTemplatePath = txtTemplatePath.Text;
            settings.LastSourceDirectory = txtSourceDirectory.Text;
            settings.FlatImport = chkFlatImport.Checked;
            settings.FlatExport = chkFlatExport.Checked;
            
            settings.Save();
        }

        protected override void OnFormClosing(System.Windows.Forms.FormClosingEventArgs e)
        {
            SaveSettings();
            logger?.Info("Building Blocks Manager closing");
            base.OnFormClosing(e);
        }

        private string ConvertCategoryToPath(string category)
        {
            // Convert "InternalAutotext\Legal\Contracts" to "Legal\Contracts"
            if (category.StartsWith("InternalAutotext\\"))
            {
                var path = category.Substring("InternalAutotext\\".Length);
                // Convert underscores back to spaces in folder names
                return path.Replace('_', ' ');
            }
            return category.Replace('_', ' ');
        }

        private string GetUniqueFilePath(string originalPath)
        {
            if (!File.Exists(originalPath))
                return originalPath;

            var directory = Path.GetDirectoryName(originalPath);
            var fileName = Path.GetFileNameWithoutExtension(originalPath);
            var extension = Path.GetExtension(originalPath);
            
            int counter = 2;
            string newPath;
            
            do
            {
                newPath = Path.Combine(directory, $"{fileName}_{counter}{extension}");
                counter++;
            }
            while (File.Exists(newPath));
            
            return newPath;
        }

        private void ViewLogMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var logDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BuildingBlocksManager", "Logs");
                
                if (Directory.Exists(logDirectory))
                {
                    var logFiles = Directory.GetFiles(logDirectory, "BBM_*.log")
                        .OrderByDescending(f => new FileInfo(f).CreationTime)
                        .FirstOrDefault();
                    
                    if (logFiles != null)
                    {
                        System.Diagnostics.Process.Start("notepad.exe", logFiles);
                    }
                    else
                    {
                        MessageBox.Show("No log files found.", "Information", 
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Log directory does not exist.", "Information", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open log file: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnLoadTemplate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTemplatePath.Text))
            {
                MessageBox.Show("Please select a template file first.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(txtTemplatePath.Text))
            {
                MessageBox.Show("The selected template file does not exist.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            UpdateStatus("Loading Building Blocks from template...");
            progressBar.Style = ProgressBarStyle.Marquee;

            try
            {
                using (var wordManager = new WordManager(txtTemplatePath.Text))
                {
                    var buildingBlocks = wordManager.GetBuildingBlocks();
                    
                    // Clear previous data
                    listTemplate.Items.Clear();
                    cmbCategoryFilter.Items.Clear();
                    
                    if (buildingBlocks.Count == 0)
                    {
                        UpdateStatus("No Building Blocks found in template");
                        return;
                    }

                    // Extract unique categories
                    var categories = buildingBlocks
                        .Select(bb => bb.Category)
                        .Where(c => !string.IsNullOrEmpty(c))
                        .Distinct()
                        .OrderBy(c => c)
                        .ToList();

                    // Populate category filter
                    cmbCategoryFilter.Items.Add("All Categories");
                    foreach (var category in categories)
                    {
                        cmbCategoryFilter.Items.Add(category);
                    }
                    cmbCategoryFilter.SelectedIndex = 0;

                    // Store all building blocks for filtering
                    listTemplate.Tag = buildingBlocks;

                    // Populate list with all items initially
                    PopulateTemplateList(buildingBlocks);

                    // Switch to Template tab
                    tabControl.SelectedTab = tabTemplate;

                    UpdateStatus($"Loaded {buildingBlocks.Count} Building Blocks from template");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus("Failed to load Building Blocks");
                MessageBox.Show($"Failed to load Building Blocks from template: {ex.Message}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
            }
        }

        private void CmbCategoryFilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listTemplate.Tag is System.Collections.Generic.List<BuildingBlockInfo> allBlocks)
            {
                var selectedCategory = cmbCategoryFilter.SelectedItem?.ToString();
                
                if (selectedCategory == "All Categories")
                {
                    PopulateTemplateList(allBlocks);
                }
                else
                {
                    var filteredBlocks = allBlocks.Where(bb => bb.Category == selectedCategory).ToList();
                    PopulateTemplateList(filteredBlocks);
                }
            }
        }

        private void PopulateTemplateList(System.Collections.Generic.List<BuildingBlockInfo> buildingBlocks)
        {
            listTemplate.Items.Clear();
            
            foreach (var bb in buildingBlocks.OrderBy(bb => bb.Name))
            {
                listTemplate.Items.Add($"{bb.Category} - {bb.Name}");
            }
        }

        private void PopulateDirectoryTree(System.Collections.Generic.List<FileManager.FileInfo> files)
        {
            treeDirectory.Nodes.Clear();
            
            // Create root node
            var rootNode = new TreeNode(Path.GetFileName(txtSourceDirectory.Text))
            {
                Tag = txtSourceDirectory.Text
            };
            treeDirectory.Nodes.Add(rootNode);
            
            // Group files by directory
            var directoryGroups = files.GroupBy(f => Path.GetDirectoryName(f.FilePath))
                                      .OrderBy(g => g.Key);
            
            foreach (var group in directoryGroups)
            {
                var dirPath = group.Key;
                var relativePath = Path.GetRelativePath(txtSourceDirectory.Text, dirPath);
                
                TreeNode parentNode = rootNode;
                
                // Create nested folder structure
                if (relativePath != ".")
                {
                    var pathParts = relativePath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    foreach (var part in pathParts)
                    {
                        var existingNode = parentNode.Nodes.Cast<TreeNode>()
                            .FirstOrDefault(n => n.Text == part && n.Tag is string);
                        
                        if (existingNode == null)
                        {
                            existingNode = new TreeNode(part)
                            {
                                Tag = Path.Combine((string)parentNode.Tag, part)
                            };
                            parentNode.Nodes.Add(existingNode);
                        }
                        
                        parentNode = existingNode;
                    }
                }
                
                // Add files to the appropriate folder node
                foreach (var file in group.OrderBy(f => Path.GetFileName(f.FilePath)))
                {
                    var fileName = Path.GetFileName(file.FilePath);
                    var status = file.IsNew ? " (New)" :
                                file.IsModified ? " (Modified)" :
                                file.IsValid ? " (Up-to-date)" : " (Invalid)";
                    
                    var fileNode = new TreeNode($"{fileName}{status}")
                    {
                        Tag = file
                    };
                    
                    // Color code the nodes
                    if (!file.IsValid)
                        fileNode.ForeColor = System.Drawing.Color.Red;
                    else if (file.IsNew)
                        fileNode.ForeColor = System.Drawing.Color.Green;
                    else if (file.IsModified)
                        fileNode.ForeColor = System.Drawing.Color.Blue;
                    
                    parentNode.Nodes.Add(fileNode);
                }
            }
            
            // Expand root node
            rootNode.Expand();
        }
    }
}