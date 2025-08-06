using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace BuildingBlocksManager
{
    public partial class MainForm : Form
    {
        private TextBox txtTemplatePath;
        private TextBox txtSourceDirectory;
        private CheckBox chkFlatImport;
        private CheckBox chkFlatExport;
        private ComboBox cmbTargetGallery;
        private Label lblTargetGallery;
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
        private ListView listViewTemplate;
        private Button btnFilterTemplate;
        private Label lblTemplateCount;
        private CheckBox chkTemplateTextsOnly;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Settings settings;
        private Logger logger;
        
        // Template filtering fields
        private List<BuildingBlockInfo> allBuildingBlocks = new List<BuildingBlockInfo>();
        private List<string> selectedCategories = new List<string>();
        private List<string> selectedGalleries = new List<string>();
        private int sortColumn = -1;
        private SortOrder sortOrder = SortOrder.None;

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
            
            var helpMenu = new ToolStripMenuItem("Help");
            var importRulesMenuItem = new ToolStripMenuItem("Import/Export Rules");
            importRulesMenuItem.Click += ImportRulesMenuItem_Click;
            helpMenu.DropDownItems.Add(importRulesMenuItem);
            
            menuStrip.Items.Add(fileMenu);
            menuStrip.Items.Add(helpMenu);
            
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

            // Gallery selection section
            lblTargetGallery = new Label
            {
                Text = "Import to Gallery:",
                Location = new System.Drawing.Point(470, 125),
                Size = new System.Drawing.Size(110, 23)
            };

            cmbTargetGallery = new ComboBox
            {
                Location = new System.Drawing.Point(590, 125),
                Size = new System.Drawing.Size(130, 23),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            
            // Populate gallery options
            cmbTargetGallery.Items.Add("AutoText");
            cmbTargetGallery.Items.Add("Quick Parts");
            cmbTargetGallery.Items.Add("Custom Gallery 1");
            cmbTargetGallery.Items.Add("Custom Gallery 2");
            cmbTargetGallery.Items.Add("Custom Gallery 3");
            cmbTargetGallery.Items.Add("Custom Gallery 4");
            cmbTargetGallery.Items.Add("Custom Gallery 5");
            cmbTargetGallery.SelectedIndex = 0; // Default to AutoText

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
                Text = "Directory",
                Location = new System.Drawing.Point(20, 190),
                Size = new System.Drawing.Size(100, 30)
            };

            var btnQueryTemplate = new Button
            {
                Text = "Template",
                Location = new System.Drawing.Point(20, 225),
                Size = new System.Drawing.Size(100, 30)
            };
            btnQueryTemplate.Click += BtnQueryTemplate_Click;

            // Import group
            var lblImport = new Label
            {
                Text = "Import\n(Folder -> Template)",
                Location = new System.Drawing.Point(160, 155),
                Size = new System.Drawing.Size(140, 35),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            btnImportAll = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(190, 195),
                Size = new System.Drawing.Size(80, 30)
            };

            btnImportSelected = new Button
            {
                Text = "Selected",
                Location = new System.Drawing.Point(190, 230),
                Size = new System.Drawing.Size(80, 30)
            };

            // Export group
            var lblExport = new Label
            {
                Text = "Export\n(Template -> Folder)",
                Location = new System.Drawing.Point(310, 155),
                Size = new System.Drawing.Size(140, 35),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            btnExportAll = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(340, 195),
                Size = new System.Drawing.Size(80, 30)
            };

            btnExportSelected = new Button
            {
                Text = "Selected",
                Location = new System.Drawing.Point(340, 230),
                Size = new System.Drawing.Size(80, 30)
            };

            // Tab control section - Form width 800px - 40px margins = 760px max
            tabControl = new TabControl
            {
                Location = new System.Drawing.Point(20, 275),
                Size = new System.Drawing.Size(740, 220)
            };

            // Results tab
            tabResults = new TabPage("Results");
            txtResults = new TextBox
            {
                Location = new System.Drawing.Point(8, 8),
                Size = new System.Drawing.Size(680, 160), // Even smaller to guarantee scrollbar space
                Multiline = true,
                ScrollBars = ScrollBars.Both, // Enable both horizontal and vertical scrollbars
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9),
                WordWrap = false // Prevent word wrapping which can cause display issues
            };
            tabResults.Controls.Add(txtResults);

            // Directory tab
            tabDirectory = new TabPage("Directory");
            treeDirectory = new TreeView
            {
                Location = new System.Drawing.Point(8, 8),
                Size = new System.Drawing.Size(680, 160), // Even smaller to guarantee scrollbar space
                Scrollable = true,
                HotTracking = true,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true
            };
            tabDirectory.Controls.Add(treeDirectory);

            // Template tab
            tabTemplate = new TabPage("Template");
            
            // Filter button
            btnFilterTemplate = new Button
            {
                Text = "Filter",
                Location = new System.Drawing.Point(5, 5),
                Size = new System.Drawing.Size(80, 25)
            };
            btnFilterTemplate.Click += BtnFilterTemplate_Click;
            
            // Item count label
            lblTemplateCount = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(95, 9),
                Size = new System.Drawing.Size(200, 17),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            
            // Template texts only checkbox
            chkTemplateTextsOnly = new CheckBox
            {
                Text = "Template texts only",
                Location = new System.Drawing.Point(305, 7),
                Size = new System.Drawing.Size(140, 20),
                Checked = false
            };
            chkTemplateTextsOnly.CheckedChanged += ChkTemplateTextsOnly_CheckedChanged;
            
            // ListView (moved down to accommodate filter controls)
            listViewTemplate = new ListView
            {
                Location = new System.Drawing.Point(8, 35),
                Size = new System.Drawing.Size(680, 130), // Even smaller to guarantee scrollbar space
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Sorting = SortOrder.None,
                Scrollable = true // Explicitly enable scrolling
            };

            // Add columns like Building Block Organizer (reduced widths to leave room for scrollbars)
            listViewTemplate.Columns.Add("Name", 220);
            listViewTemplate.Columns.Add("Category", 220);
            listViewTemplate.Columns.Add("Gallery", 180);
            
            // Enable column sorting
            listViewTemplate.ColumnClick += ListViewTemplate_ColumnClick;

            tabTemplate.Controls.Add(btnFilterTemplate);
            tabTemplate.Controls.Add(lblTemplateCount);
            tabTemplate.Controls.Add(chkTemplateTextsOnly);
            tabTemplate.Controls.Add(listViewTemplate);

            // Add tabs to tab control
            tabControl.TabPages.AddRange(new TabPage[] { tabResults, tabDirectory, tabTemplate });

            // Progress and status section
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 505),
                Size = new System.Drawing.Size(520, 23),
                Style = ProgressBarStyle.Continuous
            };

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(550, 505),
                Size = new System.Drawing.Size(210, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            // Add all controls to form
            Controls.AddRange(new Control[]
            {
                lblTemplate, txtTemplatePath, btnBrowseTemplate,
                lblDirectory, txtSourceDirectory, btnBrowseDirectory,
                lblStructure, chkFlatImport, chkFlatExport,
                lblTargetGallery, cmbTargetGallery,
                lblQuery, btnQueryDirectory, btnQueryTemplate,
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
                // Check if template file is locked and handle it
                if (!HandleTemplateFileLock(txtTemplatePath.Text, "Import All"))
                    return;

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
                        
                        // Import the Building Block with selected gallery
                        var selectedGallery = cmbTargetGallery.SelectedItem?.ToString() ?? "AutoText";
                        wordManager.ImportBuildingBlock(file.FilePath, category, file.Name, selectedGallery);
                        
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
            catch (IOException ex) when (ex.Message.Contains("locked") || ex.Message.Contains("in use"))
            {
                AppendResults($"File access error: {ex.Message}");
                AppendResults("This usually means:");
                AppendResults("• The template file is open in Word");
                AppendResults("• A previous Word process is still running");
                AppendResults("• The file is being used by another application");
                AppendResults("");
                AppendResults("Try closing Word completely and running the operation again.");
                UpdateStatus("Import failed - file access error");
                return;
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

                        // Import the Building Block with selected gallery
                        var selectedGallery = cmbTargetGallery.SelectedItem?.ToString() ?? "AutoText";
                        wordManager.ImportBuildingBlock(dialog.FileName, category, name, selectedGallery);
                        
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

            // Always prompt for export directory
            string exportPath;
            using (var dialog = new FolderBrowserDialog())
            {
                if (chkFlatExport.Checked)
                {
                    dialog.Description = "Select Export Folder (Flat Structure - All files in one folder)";
                }
                else
                {
                    dialog.Description = "Select Export Folder (Hierarchical Structure - Files organized in category folders)";
                }
                
                dialog.ShowNewFolderButton = true;
                dialog.SelectedPath = txtSourceDirectory.Text; // Default to source directory
                
                if (dialog.ShowDialog() != DialogResult.OK) return;
                exportPath = dialog.SelectedPath;
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
                // Check if template file is locked and handle it
                if (!HandleTemplateFileLock(txtTemplatePath.Text, "Export"))
                    return;

                // Initialize WordManager
                wordManager = new WordManager(txtTemplatePath.Text);
                
                // Get all Building Blocks from template
                AppendResults("Loading Building Blocks from template...");
                var allBuildingBlocks = wordManager.GetBuildingBlocks();
                
                // Filter out system/hex entries
                var buildingBlocks = allBuildingBlocks.Where(bb => !IsSystemEntry(bb)).ToList();

                if (buildingBlocks.Count == 0)
                {
                    AppendResults("No exportable Building Blocks found in template (only system entries found).");
                    return;
                }

                AppendResults($"Found {allBuildingBlocks.Count} total Building Blocks, {buildingBlocks.Count} exportable (excluding system/hex entries)");
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
                            
                            if (string.IsNullOrWhiteSpace(relativePath))
                            {
                                // No subfolder - put directly in export path
                                outputFilePath = Path.Combine(exportPath, $"AT_{bb.Name}.docx");
                            }
                            else
                            {
                                var fullDir = Path.Combine(exportPath, relativePath);
                                
                                if (!Directory.Exists(fullDir))
                                {
                                    Directory.CreateDirectory(fullDir);
                                    directoriesCreated++;
                                }
                                
                                outputFilePath = Path.Combine(fullDir, $"AT_{bb.Name}.docx");
                            }
                        }

                        // Handle duplicate filenames
                        outputFilePath = GetUniqueFilePath(outputFilePath);

                        // Export the Building Block
                        wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputFilePath);
                        
                        successCount++;
                        var displayPath = GetRelativePath(exportPath, outputFilePath);
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

            // Always prompt for export directory
            string exportPath;
            using (var dialog = new FolderBrowserDialog())
            {
                if (chkFlatExport.Checked)
                {
                    dialog.Description = "Select Export Folder (Flat Structure - All files in one folder)";
                }
                else
                {
                    dialog.Description = "Select Export Folder (Hierarchical Structure - Files organized in category folders)";
                }
                
                dialog.ShowNewFolderButton = true;
                dialog.SelectedPath = txtSourceDirectory.Text; // Default to source directory
                
                if (dialog.ShowDialog() != DialogResult.OK) return;
                exportPath = dialog.SelectedPath;
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
                // Check if template file is locked and handle it
                if (!HandleTemplateFileLock(txtTemplatePath.Text, "Export"))
                    return;

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
                            
                            if (string.IsNullOrWhiteSpace(relativePath))
                            {
                                // No subfolder - put directly in export path
                                outputFilePath = Path.Combine(exportPath, $"AT_{bb.Name}.docx");
                            }
                            else
                            {
                                var fullDir = Path.Combine(exportPath, relativePath);
                                
                                if (!Directory.Exists(fullDir))
                                {
                                    Directory.CreateDirectory(fullDir);
                                    directoriesCreated++;
                                }
                                
                                outputFilePath = Path.Combine(fullDir, $"AT_{bb.Name}.docx");
                            }
                        }

                        // Handle duplicate filenames
                        outputFilePath = GetUniqueFilePath(outputFilePath);

                        // Export the Building Block
                        wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputFilePath);
                        
                        successCount++;
                        var displayPath = GetRelativePath(exportPath, outputFilePath);
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
            // Use already loaded building blocks if available
            System.Collections.Generic.List<BuildingBlockInfo> availableBlocks;
            
            if (allBuildingBlocks.Count > 0)
            {
                availableBlocks = allBuildingBlocks.Where(bb => !IsSystemEntry(bb)).ToList();
            }
            else
            {
                // Fallback to loading from template if not already loaded
                try
                {
                    using (var wordManager = new WordManager(txtTemplatePath.Text))
                    {
                        var allBlocks = wordManager.GetBuildingBlocks();
                        availableBlocks = allBlocks.Where(bb => !IsSystemEntry(bb)).ToList();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to load Building Blocks from template: {ex.Message}", 
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return new System.Collections.Generic.List<BuildingBlockInfo>();
                }
            }

            if (availableBlocks.Count == 0)
            {
                MessageBox.Show("No exportable Building Blocks found in the template (only system/hex entries found).", 
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
            if (string.IsNullOrWhiteSpace(category))
                return ""; // No subfolder for empty categories
                
            if (category.StartsWith("InternalAutotext\\"))
            {
                var path = category.Substring("InternalAutotext\\".Length);
                // Convert underscores back to spaces in folder names
                return string.IsNullOrWhiteSpace(path) ? "" : path.Replace('_', ' ');
            }
            
            // For categories that don't follow the InternalAutotext pattern,
            // skip creating subfolders to avoid unwanted "General" folders
            return "";
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

        private string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) return toPath;
            if (string.IsNullOrEmpty(toPath)) return string.Empty;

            Uri fromUri = new Uri(AppendDirectorySeparatorChar(fromPath));
            Uri toUri = new Uri(toPath);

            if (fromUri.Scheme != toUri.Scheme) return toPath;

            Uri relativeUri = fromUri.MakeRelativeUri(toUri);
            string relativePath = Uri.UnescapeDataString(relativeUri.ToString());

            if (string.Equals(toUri.Scheme, Uri.UriSchemeFile, StringComparison.OrdinalIgnoreCase))
            {
                relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
            }

            return relativePath;
        }

        private string AppendDirectorySeparatorChar(string path)
        {
            if (!Path.HasExtension(path) && !path.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                return path + Path.DirectorySeparatorChar;
            }
            return path;
        }

        private bool HandleTemplateFileLock(string templatePath, string operationName)
        {
            if (!WordManager.IsTemplateFileLocked(templatePath))
                return true; // File is not locked, proceed

            var wordProcesses = WordManager.GetWordProcessesUsingFile(templatePath);
            
            if (wordProcesses.Count > 0)
            {
                string message = $"The template file is locked by {wordProcesses.Count} Word process(es):\n\n";
                foreach (var process in wordProcesses.Take(3))
                {
                    try
                    {
                        message += $"• Word (PID: {process.Id})\n";
                    }
                    catch
                    {
                        message += "• Word (process info unavailable)\n";
                    }
                }
                if (wordProcesses.Count > 3)
                {
                    message += $"• ... and {wordProcesses.Count - 3} more\n";
                }
                
                message += $"\nTo proceed with {operationName}, the Word processes must be closed.\n\n";
                message += "Do you want to force close these Word processes?\n\n";
                message += "WARNING: This will close Word and any unsaved documents will be lost.";
                
                var result = MessageBox.Show(
                    message,
                    "Template File Locked - Force Close Word?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                
                if (result == DialogResult.Yes) // Kill processes
                {
                    AppendResults("Force closing Word processes...");
                    bool success = WordManager.KillWordProcesses(wordProcesses);
                    
                    if (success)
                    {
                        AppendResults("Word processes closed successfully.");
                        // Wait a moment for file handles to be released
                        System.Threading.Thread.Sleep(1000);
                        
                        // Check if file is still locked
                        if (WordManager.IsTemplateFileLocked(templatePath))
                        {
                            AppendResults("Warning: File may still be locked by another process.");
                        }
                        return true; // Proceed with operation
                    }
                    else
                    {
                        AppendResults("Error: Some Word processes could not be closed. Operation cancelled.");
                        return false; // Abort if we couldn't kill processes
                    }
                }
                else // No = Cancel
                {
                    return false; // Abort operation
                }
            }
            else
            {
                var result = MessageBox.Show(
                    "The template file appears to be locked, but no Word processes were found.\n\nThis might be caused by another application or a file system issue.\n\nDo you want to continue anyway?",
                    "Template File Locked",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                    
                if (result == DialogResult.Yes)
                {
                    AppendResults("Continuing with locked file (operation may fail)...");
                    return true;
                }
                return false; // Abort operation
            }
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

        private void ImportRulesMenuItem_Click(object sender, EventArgs e)
        {
            var helpMessage = @"IMPORT/EXPORT RULES

FILE TO CATEGORY CONVERSION:
• Files must start with 'AT_' to be processed
• Directory structure converts to Building Block categories
• Top-level source directory is ignored in category path

EXAMPLE:
File: C:\MyDocs\Legal\Contracts\AT_Standard_Agreement.docx
→ Category: InternalAutotext\Legal\Contracts  
→ Name: Standard_Agreement
→ Result: Building Block appears as 'Standard_Agreement' in category 'InternalAutotext\Legal\Contracts'

SPECIAL CHARACTERS:
• Spaces in folder names become underscores in categories
• Invalid filename characters are flagged but processing continues
• Case is preserved in both folder names and Building Block names

GALLERY SELECTION:
• Choose target gallery: AutoText, Quick Parts, or Custom Gallery 1-5
• AutoText is the most commonly used gallery (default selection)
• Gallery selection affects where Building Blocks appear in Word's interface

FLAT STRUCTURE OPTIONS:
• Flat Import: All files go into single specified category
• Flat Export: All Building Blocks export to single folder (no subfolders)

BACKUP INFORMATION:
• Backups are automatically created before each import operation
• Location: Same folder as your template file
• Format: [TemplateName]_Backup_YYYYMMDD_HHMMSS.dotm
• Example: MyTemplate_Backup_20241208_143022.dotm
• Last 5 backups are kept, older ones are deleted automatically

IMPORTANT NOTES:
• Only .docx files starting with 'AT_' are processed
• Building Blocks are created in the selected gallery within the template document
• Export recreates the original folder structure from categories";

            MessageBox.Show(helpMessage, "Import/Export Rules", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnQueryTemplate_Click(object sender, EventArgs e)
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
                // Check if template file is locked and handle it
                if (!HandleTemplateFileLock(txtTemplatePath.Text, "Query Template"))
                {
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.Value = 0;
                    UpdateStatus("Query cancelled");
                    return;
                }

                using (var wordManager = new WordManager(txtTemplatePath.Text))
                {
                    var buildingBlocks = wordManager.GetBuildingBlocks();
                    
                    // Store all building blocks for filtering
                    allBuildingBlocks = buildingBlocks;
                    
                    // Reset filters when loading new template and set System/Hex to unchecked by default
                    selectedCategories.Clear();
                    selectedGalleries.Clear();
                    
                    // Apply default filter: exclude System/Hex Entries if they exist
                    var hasSystemEntries = buildingBlocks.Any(bb => IsSystemEntry(bb));
                    if (hasSystemEntries)
                    {
                        // Add all categories except System/Hex Entries to selected list
                        var categories = GetUniqueCategories();
                        foreach (var category in categories.Where(c => c != "System/Hex Entries"))
                        {
                            selectedCategories.Add(category);
                        }
                    }
                    
                    UpdateFilterButtonText();
                    
                    if (buildingBlocks.Count == 0)
                    {
                        // Clear previous data
                        listViewTemplate.Items.Clear();
                        lblTemplateCount.Text = "No Building Blocks found";
                        
                        UpdateStatus("No Building Blocks found in template");
                        tabControl.SelectedTab = tabTemplate;
                        MessageBox.Show("No Building Blocks found in the template.\n\nThis could mean:\n• The template has no Building Blocks\n• The Building Blocks are not in the 'InternalAutotext' category", 
                            "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // Apply current filter (which will exclude System/Hex entries by default)
                    ApplyTemplateFilter();

                    // Switch to Template tab
                    tabControl.SelectedTab = tabTemplate;

                    UpdateStatus($"Loaded {buildingBlocks.Count} Building Blocks from template");
                }
            }
            catch (UnauthorizedAccessException)
            {
                UpdateStatus("Template file access denied");
                MessageBox.Show("Cannot access the template file. Make sure:\n• The file is not open in Word\n• You have permission to read the file", 
                    "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (IOException ex)
            {
                UpdateStatus("Template file I/O error");
                MessageBox.Show($"File I/O error: {ex.Message}\n\nMake sure the template file is not corrupted and not open in another application.", 
                    "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (COMException ex)
            {
                UpdateStatus("Word COM error");
                MessageBox.Show($"Word COM error: {ex.Message}\n\nThis could mean:\n• Word is not installed or not working properly\n• The template file is corrupted\n• Word COM components need repair", 
                    "Word Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                UpdateStatus("Failed to load Building Blocks");
                MessageBox.Show($"Unexpected error: {ex.Message}\n\nFull details:\n{ex}", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
            }
        }


        private void PopulateDirectoryTree(System.Collections.Generic.List<FileManager.FileInfo> files)
        {
            treeDirectory.Nodes.Clear();
            
            if (!Directory.Exists(txtSourceDirectory.Text))
                return;
                
            // Use standard DirectoryInfo approach
            var rootDir = new DirectoryInfo(txtSourceDirectory.Text);
            var rootNode = CreateDirectoryNode(rootDir, files);
            
            treeDirectory.Nodes.Add(rootNode);
            rootNode.Expand();
        }
        
        private TreeNode CreateDirectoryNode(DirectoryInfo directory, System.Collections.Generic.List<FileManager.FileInfo> scannedFiles)
        {
            var node = new TreeNode(directory.Name)
            {
                Tag = directory.FullName
            };
            
            try
            {
                // Add subdirectories
                var subdirs = directory.GetDirectories().OrderBy(d => d.Name);
                foreach (var subdir in subdirs)
                {
                    var childNode = CreateDirectoryNode(subdir, scannedFiles);
                    node.Nodes.Add(childNode);
                }
                
                // Add files - only show AT_ files that were scanned
                var relevantFiles = scannedFiles.Where(f => 
                    Path.GetDirectoryName(f.FilePath).Equals(directory.FullName, StringComparison.OrdinalIgnoreCase))
                    .OrderBy(f => Path.GetFileName(f.FilePath));
                
                foreach (var file in relevantFiles)
                {
                    var fileName = Path.GetFileName(file.FilePath);
                    var status = file.IsNew ? " (New)" :
                                file.IsModified ? " (Modified)" :
                                file.IsValid ? " (Up-to-date)" : " (Invalid)";
                    
                    var fileNode = new TreeNode($"{fileName}{status}")
                    {
                        Tag = file
                    };
                    
                    // Color code the file nodes
                    if (!file.IsValid)
                        fileNode.ForeColor = System.Drawing.Color.Red;
                    else if (file.IsNew)
                        fileNode.ForeColor = System.Drawing.Color.Green;
                    else if (file.IsModified)
                        fileNode.ForeColor = System.Drawing.Color.Blue;
                    else
                        fileNode.ForeColor = System.Drawing.Color.Black;
                    
                    node.Nodes.Add(fileNode);
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Skip directories we can't access
                node.Text += " (Access Denied)";
                node.ForeColor = System.Drawing.Color.Gray;
            }
            
            return node;
        }

        private void BtnFilterTemplate_Click(object sender, EventArgs e)
        {
            if (allBuildingBlocks.Count == 0)
            {
                MessageBox.Show("Please load building blocks from template first.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var filterDialog = CreateFilterDialog();
            if (filterDialog.ShowDialog() == DialogResult.OK)
            {
                ApplyTemplateFilter();
                UpdateFilterButtonText();
            }
            filterDialog.Dispose();
        }

        private Form CreateFilterDialog()
        {
            var categories = GetUniqueCategories();
            var galleries = GetUniqueGalleries();
            
            var form = new Form
            {
                Text = "Filter Building Blocks",
                Size = new System.Drawing.Size(600, 500),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            // Categories section
            var lblCategories = new Label
            {
                Text = "Categories:",
                Location = new System.Drawing.Point(20, 10),
                Size = new System.Drawing.Size(100, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            var listBoxCategories = new CheckedListBox
            {
                Location = new System.Drawing.Point(20, 35),
                Size = new System.Drawing.Size(250, 280),
                CheckOnClick = true
            };

            foreach (var category in categories)
            {
                bool isChecked;
                if (category == "System/Hex Entries")
                {
                    // System/Hex Entries should default to unchecked
                    isChecked = selectedCategories.Contains(category);
                }
                else
                {
                    // Other categories default to checked if no filters are applied
                    isChecked = selectedCategories.Count == 0 || selectedCategories.Contains(category);
                }
                listBoxCategories.Items.Add(category, isChecked);
            }

            // Galleries section
            var lblGalleries = new Label
            {
                Text = "Galleries:",
                Location = new System.Drawing.Point(300, 10),
                Size = new System.Drawing.Size(100, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            var listBoxGalleries = new CheckedListBox
            {
                Location = new System.Drawing.Point(300, 35),
                Size = new System.Drawing.Size(250, 280),
                CheckOnClick = true
            };

            foreach (var gallery in galleries)
            {
                bool isChecked = selectedGalleries.Count == 0 || selectedGalleries.Contains(gallery);
                listBoxGalleries.Items.Add(gallery, isChecked);
            }

            var btnSelectAllCat = new Button
            {
                Text = "All Categories",
                Location = new System.Drawing.Point(20, 330),
                Size = new System.Drawing.Size(100, 25)
            };

            var btnSelectNoneCat = new Button
            {
                Text = "No Categories",
                Location = new System.Drawing.Point(130, 330),
                Size = new System.Drawing.Size(100, 25)
            };

            var btnSelectAllGal = new Button
            {
                Text = "All Galleries",
                Location = new System.Drawing.Point(300, 330),
                Size = new System.Drawing.Size(100, 25)
            };

            var btnSelectNoneGal = new Button
            {
                Text = "No Galleries",
                Location = new System.Drawing.Point(410, 330),
                Size = new System.Drawing.Size(100, 25)
            };

            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(400, 420),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(490, 420),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.Cancel
            };

            btnSelectAllCat.Click += (s, e) => {
                for (int i = 0; i < listBoxCategories.Items.Count; i++)
                    listBoxCategories.SetItemChecked(i, true);
            };

            btnSelectNoneCat.Click += (s, e) => {
                for (int i = 0; i < listBoxCategories.Items.Count; i++)
                    listBoxCategories.SetItemChecked(i, false);
            };

            btnSelectAllGal.Click += (s, e) => {
                for (int i = 0; i < listBoxGalleries.Items.Count; i++)
                    listBoxGalleries.SetItemChecked(i, true);
            };

            btnSelectNoneGal.Click += (s, e) => {
                for (int i = 0; i < listBoxGalleries.Items.Count; i++)
                    listBoxGalleries.SetItemChecked(i, false);
            };

            btnOK.Click += (s, e) => {
                selectedCategories.Clear();
                foreach (string item in listBoxCategories.CheckedItems)
                {
                    selectedCategories.Add(item);
                }
                selectedGalleries.Clear();
                foreach (string item in listBoxGalleries.CheckedItems)
                {
                    selectedGalleries.Add(item);
                }
            };

            form.Controls.AddRange(new Control[] { lblCategories, listBoxCategories, lblGalleries, listBoxGalleries, 
                btnSelectAllCat, btnSelectNoneCat, btnSelectAllGal, btnSelectNoneGal, btnOK, btnCancel });
            form.AcceptButton = btnOK;
            form.CancelButton = btnCancel;

            return form;
        }

        private List<string> GetUniqueCategories()
        {
            var categories = allBuildingBlocks
                .Select(bb => string.IsNullOrEmpty(bb.Category) ? "(No Category)" : bb.Category)
                .Distinct()
                .OrderBy(cat => cat)
                .ToList();

            // Add "System/Hex Entries" as a special category for filtering system entries
            if (allBuildingBlocks.Any(bb => IsSystemEntry(bb)))
            {
                categories.Insert(0, "System/Hex Entries");
            }

            return categories;
        }

        private List<string> GetUniqueGalleries()
        {
            var galleries = allBuildingBlocks
                .Select(bb => string.IsNullOrEmpty(bb.Gallery) ? "(No Gallery)" : bb.Gallery)
                .Distinct()
                .OrderBy(gal => gal)
                .ToList();

            return galleries;
        }

        private bool IsSystemEntry(BuildingBlockInfo bb)
        {
            // More targeted system entry detection
            return bb.Name.Length > 15 && bb.Name.All(c => "0123456789ABCDEF-".Contains(char.ToUpper(c))) ||
                   bb.Name.StartsWith("_") ||
                   (bb.Category != null && bb.Category.Contains("System"));
        }

        private void ApplyTemplateFilter()
        {
            listViewTemplate.Items.Clear();
            
            List<BuildingBlockInfo> filteredBlocks;
            
            if (selectedCategories.Count == 0 && selectedGalleries.Count == 0)
            {
                // No filter - show all
                filteredBlocks = allBuildingBlocks;
            }
            else
            {
                filteredBlocks = new List<BuildingBlockInfo>();
                
                foreach (var bb in allBuildingBlocks)
                {
                    bool includeByCategory = selectedCategories.Count == 0;
                    bool includeByGallery = selectedGalleries.Count == 0;
                    bool includeByTemplateOnly = true;
                    
                    // Check "Template texts only" filter
                    if (chkTemplateTextsOnly.Checked)
                    {
                        // Only include AutoText gallery items (exclude Built-In, etc.)
                        includeByTemplateOnly = bb.Gallery == "AutoText";
                    }
                    
                    // Check category filter
                    if (selectedCategories.Count > 0)
                    {
                        if (selectedCategories.Contains("System/Hex Entries") && IsSystemEntry(bb))
                        {
                            includeByCategory = true;
                        }
                        else if (selectedCategories.Contains(string.IsNullOrEmpty(bb.Category) ? "(No Category)" : bb.Category))
                        {
                            if (selectedCategories.Contains("System/Hex Entries") || !IsSystemEntry(bb))
                            {
                                includeByCategory = true;
                            }
                        }
                    }
                    
                    // Check gallery filter
                    if (selectedGalleries.Count > 0)
                    {
                        if (selectedGalleries.Contains(string.IsNullOrEmpty(bb.Gallery) ? "(No Gallery)" : bb.Gallery))
                        {
                            includeByGallery = true;
                        }
                    }
                    
                    if (includeByCategory && includeByGallery && includeByTemplateOnly)
                    {
                        filteredBlocks.Add(bb);
                    }
                }
            }

            // Populate ListView with filtered building blocks
            foreach (var bb in filteredBlocks.OrderBy(bb => bb.Name))
            {
                var item = new ListViewItem(bb.Name);
                item.SubItems.Add(bb.Category);
                item.SubItems.Add(bb.Gallery);
                item.Tag = bb;
                listViewTemplate.Items.Add(item);
            }

            UpdateTemplateCount(filteredBlocks.Count);
        }

        private void UpdateFilterButtonText()
        {
            int totalFilters = selectedCategories.Count + selectedGalleries.Count;
            if (totalFilters == 0)
            {
                btnFilterTemplate.Text = "Filter";
            }
            else
            {
                btnFilterTemplate.Text = $"Filter: {totalFilters}";
            }
        }

        private void UpdateTemplateCount(int filteredCount)
        {
            if (filteredCount == allBuildingBlocks.Count)
            {
                lblTemplateCount.Text = $"Showing {allBuildingBlocks.Count} Building Blocks";
            }
            else
            {
                lblTemplateCount.Text = $"Showing {filteredCount} of {allBuildingBlocks.Count} Building Blocks";
            }
        }

        private void ListViewTemplate_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            if (e.Column == sortColumn)
            {
                sortOrder = sortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                sortColumn = e.Column;
                sortOrder = SortOrder.Ascending;
            }

            listViewTemplate.ListViewItemSorter = new ListViewItemComparer(e.Column, sortOrder);
            listViewTemplate.Sort();
        }

        private void ChkTemplateTextsOnly_CheckedChanged(object sender, EventArgs e)
        {
            // Re-apply current filters when template texts only checkbox changes
            ApplyTemplateFilter();
        }
    }

    public class ListViewItemComparer : System.Collections.IComparer
    {
        private int column;
        private SortOrder order;

        public ListViewItemComparer(int column, SortOrder order)
        {
            this.column = column;
            this.order = order;
        }

        public int Compare(object x, object y)
        {
            int returnVal = String.Compare(((ListViewItem)x).SubItems[column].Text, 
                                         ((ListViewItem)y).SubItems[column].Text);
            return order == SortOrder.Descending ? -returnVal : returnVal;
        }
    }
}