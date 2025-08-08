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
        private TextBox txtExportDirectory;
        private CheckBox chkFlatImport;
        private CheckBox chkFlatExport;
        private Button btnBrowseTemplate;
        private Button btnBrowseDirectory;
        private Button btnBrowseExportDirectory;
        private Label lblTemplatePathDisplay;
        private Label lblSourceDirectoryPathDisplay;
        private Label lblExportDirectoryPathDisplay;
        private Button btnQueryDirectory;
        private Button btnImportAll;
        private Button btnImportSelected;
        private Button btnExportAll;
        private Button btnExportSelected;
        private Button btnStop;
        private TextBox txtResults;
        private TabControl tabControl;
        private TabPage tabResults;
        private TabPage tabDirectory;
        private TabPage tabTemplate;
        private TreeView treeDirectory;
        private ListView listViewTemplate;
        private Button btnFilterTemplate;
        private Label lblTemplateCount;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Settings settings;
        private Logger logger;
        private System.Threading.CancellationTokenSource cancellationTokenSource;
        
        // Store the actual full paths since text boxes now show partial info
        private string fullTemplatePath = "";
        private string fullSourceDirectoryPath = "";
        private string fullExportDirectoryPath = "";
        
        // Template filtering fields
        private List<BuildingBlockInfo> allBuildingBlocks = new List<BuildingBlockInfo>();
        private List<string> selectedCategories = new List<string>();
        private List<string> selectedGalleries = new List<string>();
        private List<string> selectedTemplates = new List<string>();
        private int sortColumn = -1;
        private SortOrder sortOrder = SortOrder.None;
        private System.Collections.Generic.List<string> conflictedFiles = new System.Collections.Generic.List<string>();

        public MainForm()
        {
            InitializeComponent();
            this.Text = "Building Blocks Manager";
            this.Size = new System.Drawing.Size(600, 600);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new System.Drawing.Size(450, 500);
            
            // Wire up event handlers
            btnBrowseTemplate.Click += BtnBrowseTemplate_Click;
            btnBrowseDirectory.Click += BtnBrowseDirectory_Click;
            btnBrowseExportDirectory.Click += BtnBrowseExportDirectory_Click;
            btnQueryDirectory.Click += BtnQueryDirectory_Click;
            btnImportAll.Click += BtnImportAll_Click;
            btnImportSelected.Click += BtnImportSelected_Click;
            btnExportAll.Click += BtnExportAll_Click;
            btnExportSelected.Click += BtnExportSelected_Click;
            
            // Load and apply settings
            LoadSettings();
            
            // Initialize logger (will be reinitialized when source directory is selected)
            InitializeLogger();
        }

        private void InitializeLogger()
        {
            logger = new Logger(fullTemplatePath, fullSourceDirectoryPath, settings.LogToTemplateDirectory, settings.EnableDetailedLogging);
            logger.CleanupOldLogs();
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
            
            var loggingConfigMenuItem = new ToolStripMenuItem("Logging Configuration...");
            loggingConfigMenuItem.Click += LoggingConfigMenuItem_Click;
            fileMenu.DropDownItems.Add(loggingConfigMenuItem);
            
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
                Location = new System.Drawing.Point(10, 45),
                Size = new System.Drawing.Size(85, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            txtTemplatePath = new TextBox
            {
                Location = new System.Drawing.Point(100, 45),
                Size = new System.Drawing.Size(280, 23),
                ReadOnly = true
            };

            lblTemplatePathDisplay = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(100, 70),
                Size = new System.Drawing.Size(280, 15),
                Font = new System.Drawing.Font(Label.DefaultFont.FontFamily, Label.DefaultFont.Size - 1, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.Gray
            };

            btnBrowseTemplate = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(390, 44),
                Size = new System.Drawing.Size(70, 25)
            };

            // Source directory section
            var lblDirectory = new Label
            {
                Text = "Source:",
                Location = new System.Drawing.Point(10, 95),
                Size = new System.Drawing.Size(45, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            txtSourceDirectory = new TextBox
            {
                Location = new System.Drawing.Point(60, 95),
                Size = new System.Drawing.Size(140, 23),
                ReadOnly = true
            };

            lblSourceDirectoryPathDisplay = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(60, 120),
                Size = new System.Drawing.Size(140, 15),
                Font = new System.Drawing.Font(Label.DefaultFont.FontFamily, Label.DefaultFont.Size - 1, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.Gray
            };

            btnBrowseDirectory = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(210, 94),
                Size = new System.Drawing.Size(60, 25)
            };

            // Export directory section
            var lblExportDirectory = new Label
            {
                Text = "Export:",
                Location = new System.Drawing.Point(280, 95),
                Size = new System.Drawing.Size(45, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            txtExportDirectory = new TextBox
            {
                Location = new System.Drawing.Point(330, 95),
                Size = new System.Drawing.Size(140, 23),
                ReadOnly = true
            };

            lblExportDirectoryPathDisplay = new Label
            {
                Text = "",
                Location = new System.Drawing.Point(330, 120),
                Size = new System.Drawing.Size(140, 15),
                Font = new System.Drawing.Font(Label.DefaultFont.FontFamily, Label.DefaultFont.Size - 1, System.Drawing.FontStyle.Regular),
                ForeColor = System.Drawing.Color.Gray
            };

            btnBrowseExportDirectory = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(480, 94),
                Size = new System.Drawing.Size(60, 25)
            };

            // Structure options section
            var lblStructure = new Label
            {
                Text = "Ignore folder/category structure for:",
                Location = new System.Drawing.Point(20, 150),
                Size = new System.Drawing.Size(250, 23)
            };

            chkFlatImport = new CheckBox
            {
                Text = "Import",
                Location = new System.Drawing.Point(280, 150),
                Size = new System.Drawing.Size(80, 23)
            };

            chkFlatExport = new CheckBox
            {
                Text = "Export",
                Location = new System.Drawing.Point(370, 150),
                Size = new System.Drawing.Size(80, 23)
            };

            // Query group
            var lblQuery = new Label
            {
                Text = "Query",
                Location = new System.Drawing.Point(20, 190),
                Size = new System.Drawing.Size(100, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            btnQueryDirectory = new Button
            {
                Text = "Directory",
                Location = new System.Drawing.Point(20, 215),
                Size = new System.Drawing.Size(100, 30)
            };

            var btnQueryTemplate = new Button
            {
                Text = "Template",
                Location = new System.Drawing.Point(20, 250),
                Size = new System.Drawing.Size(100, 30)
            };
            btnQueryTemplate.Click += BtnQueryTemplate_Click;

            // Import group
            var lblImport = new Label
            {
                Text = "Import\n(Folder -> Template)",
                Location = new System.Drawing.Point(160, 180),
                Size = new System.Drawing.Size(140, 35),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            btnImportAll = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(190, 220),
                Size = new System.Drawing.Size(80, 30)
            };

            btnImportSelected = new Button
            {
                Text = "Selected",
                Location = new System.Drawing.Point(190, 255),
                Size = new System.Drawing.Size(80, 30)
            };

            // Export group
            var lblExport = new Label
            {
                Text = "Export\n(Template -> Folder)",
                Location = new System.Drawing.Point(310, 180),
                Size = new System.Drawing.Size(140, 35),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold),
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            };

            btnExportAll = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(340, 220),
                Size = new System.Drawing.Size(80, 30)
            };

            btnExportSelected = new Button
            {
                Text = "Selected",
                Location = new System.Drawing.Point(340, 255),
                Size = new System.Drawing.Size(80, 30)
            };

            // Stop button (hidden by default)
            btnStop = new Button
            {
                Text = "Stop",
                Location = new System.Drawing.Point(450, 235),
                Size = new System.Drawing.Size(80, 35),
                Visible = false,
                BackColor = System.Drawing.Color.LightCoral
            };
            btnStop.Click += BtnStop_Click;

            // Tab control section - Form width 600px - 40px margins = 560px max (25% reduction from 740)
            tabControl = new TabControl
            {
                Location = new System.Drawing.Point(20, 300),
                Size = new System.Drawing.Size(555, 220)
            };

            // Results tab
            tabResults = new TabPage("Results");
            txtResults = new TextBox
            {
                Location = new System.Drawing.Point(3, 3),
                Size = new System.Drawing.Size(540, 185), // 25% reduction from 725
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
                Location = new System.Drawing.Point(3, 3),
                Size = new System.Drawing.Size(540, 185), // 25% reduction from 725
                Scrollable = true,
                HotTracking = true,
                ShowLines = true,
                ShowPlusMinus = true,
                ShowRootLines = true,
                CheckBoxes = true // Enable checkboxes for multiselect
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
            
            // Select All button
            var btnSelectAll = new Button
            {
                Text = "Select All",
                Location = new System.Drawing.Point(305, 5),
                Size = new System.Drawing.Size(70, 25)
            };
            btnSelectAll.Click += BtnSelectAll_Click;
            
            // Delete tip label
            var lblDeleteTip = new Label
            {
                Text = "Right-click to delete autotext",
                Location = new System.Drawing.Point(385, 9),
                Size = new System.Drawing.Size(140, 17),
                Font = new System.Drawing.Font(Label.DefaultFont.FontFamily, Label.DefaultFont.Size, System.Drawing.FontStyle.Italic),
                ForeColor = System.Drawing.Color.Gray
            };
            
            // ListView (moved down to accommodate filter controls)
            listViewTemplate = new ListView
            {
                Location = new System.Drawing.Point(3, 35),
                Size = new System.Drawing.Size(540, 150), // 25% reduction from 725
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Sorting = SortOrder.None,
                Scrollable = true, // Explicitly enable scrolling
                MultiSelect = true // Enable multi-selection
            };

            // Add columns - adjusted for 25% width reduction (540px total width)
            listViewTemplate.Columns.Add("Name", 135);      // 135px (25% reduction from 180)
            listViewTemplate.Columns.Add("Gallery", 135);   // 135px (25% reduction from 180) 
            listViewTemplate.Columns.Add("Category", 135);  // 135px (25% reduction from 180)
            listViewTemplate.Columns.Add("Template", 120);  // 120px (25% reduction from 160)
            // Total: 525px, leaving 15px buffer for scrollbars
            
            // Enable column sorting
            listViewTemplate.ColumnClick += ListViewTemplate_ColumnClick;
            
            // Enable context menu for template management
            listViewTemplate.ContextMenuStrip = CreateTemplateContextMenu();
            listViewTemplate.KeyDown += ListViewTemplate_KeyDown;

            tabTemplate.Controls.Add(btnFilterTemplate);
            tabTemplate.Controls.Add(lblTemplateCount);
            tabTemplate.Controls.Add(btnSelectAll);
            tabTemplate.Controls.Add(lblDeleteTip);
            tabTemplate.Controls.Add(listViewTemplate);

            // Add tabs to tab control
            tabControl.TabPages.AddRange(new TabPage[] { tabResults, tabDirectory, tabTemplate });
            
            // Add event handler for automatic querying when tabs are first accessed
            tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;

            // Progress and status section
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 530),
                Size = new System.Drawing.Size(390, 23),
                Style = ProgressBarStyle.Continuous
            };

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(420, 530),
                Size = new System.Drawing.Size(160, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            // Add all controls to form
            Controls.AddRange(new Control[]
            {
                lblTemplate, txtTemplatePath, lblTemplatePathDisplay, btnBrowseTemplate,
                lblDirectory, txtSourceDirectory, lblSourceDirectoryPathDisplay, btnBrowseDirectory,
                lblExportDirectory, txtExportDirectory, lblExportDirectoryPathDisplay, btnBrowseExportDirectory,
                lblStructure, chkFlatImport, chkFlatExport,
                lblQuery, btnQueryDirectory, btnQueryTemplate,
                lblImport, btnImportAll, btnImportSelected,
                lblExport, btnExportAll, btnExportSelected, btnStop,
                tabControl,
                progressBar, lblStatus
            });

            ResumeLayout(false);
        }

        private void UpdateTemplatePathDisplay(string fullPath)
        {
            fullTemplatePath = fullPath ?? "";
            
            if (string.IsNullOrEmpty(fullPath))
            {
                txtTemplatePath.Text = "";
                lblTemplatePathDisplay.Text = "";
                return;
            }

            var fileName = Path.GetFileName(fullPath);
            var directoryPath = Path.GetDirectoryName(fullPath);
            
            txtTemplatePath.Text = fileName;
            lblTemplatePathDisplay.Text = string.IsNullOrEmpty(directoryPath) ? "" : directoryPath + Path.DirectorySeparatorChar;
        }

        private void UpdateSourceDirectoryDisplay(string fullPath)
        {
            fullSourceDirectoryPath = fullPath ?? "";
            
            if (string.IsNullOrEmpty(fullPath))
            {
                txtSourceDirectory.Text = "";
                lblSourceDirectoryPathDisplay.Text = "";
                return;
            }

            var directoryInfo = new DirectoryInfo(fullPath);
            var lowestLevelDirectory = directoryInfo.Name;
            var parentPath = directoryInfo.Parent?.FullName;
            
            txtSourceDirectory.Text = "..." + Path.DirectorySeparatorChar + lowestLevelDirectory;
            lblSourceDirectoryPathDisplay.Text = string.IsNullOrEmpty(parentPath) ? "" : parentPath + Path.DirectorySeparatorChar;
            
            // Reinitialize logger with new source directory
            InitializeLogger();
        }

        private void UpdateExportDirectoryDisplay(string fullPath)
        {
            fullExportDirectoryPath = fullPath ?? "";
            
            if (string.IsNullOrEmpty(fullPath))
            {
                txtExportDirectory.Text = "";
                lblExportDirectoryPathDisplay.Text = "";
                return;
            }

            var directoryInfo = new DirectoryInfo(fullPath);
            var lowestLevelDirectory = directoryInfo.Name;
            var parentPath = directoryInfo.Parent?.FullName;
            
            txtExportDirectory.Text = "..." + Path.DirectorySeparatorChar + lowestLevelDirectory;
            lblExportDirectoryPathDisplay.Text = string.IsNullOrEmpty(parentPath) ? "" : parentPath + Path.DirectorySeparatorChar;
        }

        private void TabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Auto-query when Directory or Template tabs are first accessed and empty
            if (tabControl.SelectedTab == tabDirectory)
            {
                // Check if directory tree is empty - trigger directory query
                if (treeDirectory.Nodes.Count == 0 && ValidatePaths())
                {
                    BtnQueryDirectory_Click(sender, e);
                }
            }
            else if (tabControl.SelectedTab == tabTemplate)
            {
                // Check if template list is empty - trigger template query
                if (listViewTemplate.Items.Count == 0 && !string.IsNullOrWhiteSpace(fullTemplatePath) && File.Exists(fullTemplatePath))
                {
                    BtnQueryTemplate_Click(sender, e);
                }
            }
        }

        private void BtnStop_Click(object sender, EventArgs e)
        {
            cancellationTokenSource?.Cancel();
            UpdateStatus("Operation cancelled by user");
        }

        private void ShowStopButton()
        {
            btnStop.Visible = true;
            cancellationTokenSource = new System.Threading.CancellationTokenSource();
        }

        private void HideStopButton()
        {
            btnStop.Visible = false;
            cancellationTokenSource?.Dispose();
            cancellationTokenSource = null;
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
                    UpdateTemplatePathDisplay(dialog.FileName);
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
                    UpdateSourceDirectoryDisplay(dialog.SelectedPath);
                    UpdateStatus("Source directory selected: " + dialog.SelectedPath);
                    SaveSettings();
                }
            }
        }

        private void BtnBrowseExportDirectory_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Export Directory";
                dialog.ShowNewFolderButton = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    UpdateExportDirectoryDisplay(dialog.SelectedPath);
                    UpdateStatus("Export directory selected: " + dialog.SelectedPath);
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
            txtResults.AppendText($"[{DateTime.Now:HH:mm}] {message}\r\n");
            txtResults.ScrollToCaret();
            Application.DoEvents();
        }

        private bool ValidatePaths()
        {
            if (string.IsNullOrWhiteSpace(fullTemplatePath))
            {
                MessageBox.Show("Please select a template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!File.Exists(fullTemplatePath))
            {
                MessageBox.Show("The selected template file does not exist.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(fullSourceDirectoryPath))
            {
                MessageBox.Show("Please select a source directory.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!Directory.Exists(fullSourceDirectoryPath))
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
            AppendResults($"Scanning directory: {fullSourceDirectoryPath}");
            
            try
            {
                var fileManager = new FileManager(fullSourceDirectoryPath);
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
                AppendResults($"Using flat import category: {flatCategory}");
            }

            UpdateStatus("Importing all files...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            ShowStopButton();
            
            AppendResults("=== IMPORT ALL OPERATION ===");
            AppendResults($"Template: {Path.GetFileName(fullTemplatePath)}");
            AppendResults($"Source Directory: {fullSourceDirectoryPath}");
            AppendResults($"Flat Import: {(chkFlatImport.Checked ? "Yes" : "No")}");
            AppendResults("");
            
            logger.Info($"Starting Import All operation - Template: {fullTemplatePath}, Directory: {fullSourceDirectoryPath}");

            WordManager wordManager = null;
            var startTime = DateTime.Now;
            int successCount = 0;
            int errorCount = 0;

            try
            {
                // Check if template file is locked and handle it
                if (!HandleTemplateFileLock(fullTemplatePath, "Import All"))
                {
                    HideStopButton();
                    return;
                }

                // Initialize managers
                wordManager = new WordManager(fullTemplatePath);
                var fileManager = new FileManager(fullSourceDirectoryPath);
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
                    HideStopButton();
                    return;
                }

                AppendResults($"Found {filesToImport.Count} files to import");
                AppendResults("");

                // Import each file
                for (int i = 0; i < filesToImport.Count; i++)
                {
                    // Check for cancellation
                    if (cancellationTokenSource?.Token.IsCancellationRequested == true)
                    {
                        AppendResults("Import operation cancelled by user.");
                        break;
                    }

                    var file = filesToImport[i];
                    var fileName = Path.GetFileName(file.FilePath);
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / filesToImport.Count * 100);
                        UpdateStatus($"Processing {fileName}...");
                        AppendResults($"Processing {fileName}...");

                        // Use flat category if specified, otherwise use extracted category
                        var category = chkFlatImport.Checked ? flatCategory : file.Category;
                        
                        // Import the Building Block to AutoText gallery
                        wordManager.ImportBuildingBlock(file.FilePath, category, file.Name, "AutoText");
                        
                        // Update import tracking
                        importTracker.UpdateImportTime(file.FilePath);
                        
                        successCount++;
                        AppendResults($"  ✓ Successfully imported as {category}\\{file.Name}");
                        logger.Success($"Imported {fileName} as {category}\\{file.Name}");
                        logger.LogImport(fileName, file.Name, category, file.FilePath);
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
                HideStopButton();
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
                HideStopButton();
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

            // Get checked file nodes from TreeView
            var checkedFiles = GetCheckedFiles(treeDirectory.Nodes);
            
            if (checkedFiles.Count == 0)
            {
                MessageBox.Show("Please check one or more AT_ files in the Directory tab to import.", 
                    "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            UpdateStatus("Importing selected files...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            
            AppendResults("=== IMPORT SELECTED FILES ===");
            AppendResults($"Selected Files: {checkedFiles.Count}");
            AppendResults($"Template: {Path.GetFileName(fullTemplatePath)}");
            AppendResults("");

            WordManager wordManager = null;
            var startTime = DateTime.Now;
            int successCount = 0;
            int errorCount = 0;
            
            // Check for flat import category if needed
            string flatCategory = null;
            if (chkFlatImport.Checked)
            {
                flatCategory = PromptForInput("Flat Import Category", 
                    "Enter category name for all selected Building Blocks:");
                if (string.IsNullOrWhiteSpace(flatCategory))
                {
                    AppendResults("Import cancelled - no category specified.");
                    return;
                }
            }

            try
            {
                // Initialize managers
                wordManager = new WordManager(fullTemplatePath);
                var fileManager = new FileManager(fullSourceDirectoryPath);
                var importTracker = new ImportTracker();

                // Create backup
                AppendResults("Creating backup...");
                wordManager.CreateBackup();

                // Process each selected file
                for (int i = 0; i < checkedFiles.Count; i++)
                {
                    var file = checkedFiles[i];
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / checkedFiles.Count * 100);
                        UpdateStatus($"Importing: {i + 1} of {checkedFiles.Count}");
                        
                        var fileName = Path.GetFileName(file.FilePath);
                        AppendResults($"Processing {fileName}...");

                        var category = chkFlatImport.Checked ? flatCategory : file.Category;
                        var name = fileManager.ExtractName(fileName);

                        // Check for invalid characters
                        var invalidChars = fileManager.GetInvalidCharacters(fileName);
                        if (invalidChars.Count > 0)
                        {
                            AppendResults($"  Warning: File contains invalid characters: {string.Join(", ", invalidChars)}");
                        }

                        // Import the Building Block
                        wordManager.ImportBuildingBlock(file.FilePath, category, name, "AutoText");
                        
                        // Update import tracking
                        importTracker.UpdateImportTime(file.FilePath);
                        
                        successCount++;
                        logger.Success($"Imported {fileName} as {category}\\{name}");
                        logger.LogImport(fileName, name, category, file.FilePath);
                        AppendResults($"  ✓ Imported as {category}\\{name}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        var fileName = Path.GetFileName(file.FilePath);
                        logger.Error($"Failed to import {fileName}: {ex.Message}");
                        AppendResults($"  ✗ Failed to import {fileName}: {ex.Message}");
                    }
                }
                
                var processingTime = (DateTime.Now - startTime).TotalSeconds;
                
                AppendResults("");
                AppendResults($"Import Summary: Success: {successCount}, Errors: {errorCount}");
                AppendResults($"Processing Time: {processingTime:F1} seconds");
                
                logger.Info($"Import Selected completed - Success: {successCount}, Errors: {errorCount}, Time: {processingTime:F1}s");
            }
            catch (Exception ex)
            {
                AppendResults($"Import operation failed: {ex.Message}");
                logger.Error($"Import Selected operation failed: {ex.Message}");
            }
            finally
            {
                wordManager?.Dispose();
            }
            
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            UpdateStatus(errorCount == 0 ? "Import completed successfully" : $"Import completed with {errorCount} errors");
        }

        private void BtnExportAll_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            // Use export directory if set, otherwise prompt
            string exportPath;
            if (!string.IsNullOrEmpty(fullExportDirectoryPath) && Directory.Exists(fullExportDirectoryPath))
            {
                exportPath = fullExportDirectoryPath;
            }
            else
            {
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
                    dialog.SelectedPath = !string.IsNullOrEmpty(fullExportDirectoryPath) ? fullExportDirectoryPath : fullSourceDirectoryPath;
                    
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    exportPath = dialog.SelectedPath;
                }
            }

            UpdateStatus("Exporting all Building Blocks...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            ShowStopButton();
            
            // Clear conflict list for this operation
            conflictedFiles.Clear();
            
            AppendResults("=== EXPORT ALL OPERATION ===");
            AppendResults($"Template: {Path.GetFileName(fullTemplatePath)}");
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
                if (!HandleTemplateFileLock(fullTemplatePath, "Export"))
                {
                    HideStopButton();
                    return;
                }

                // Initialize WordManager
                wordManager = new WordManager(fullTemplatePath);
                
                // Check if building blocks have been loaded (user must run Query Template first)
                if (allBuildingBlocks.Count == 0)
                {
                    AppendResults("No Building Blocks loaded. Please run Query Template first to load and filter Building Blocks.");
                    HideStopButton();
                    return;
                }
                
                // Use the current filtered list from the UI
                var buildingBlocks = GetFilteredBuildingBlocks();
                AppendResults("Using filtered Building Blocks from current display...");

                if (buildingBlocks.Count == 0)
                {
                    AppendResults("No Building Blocks to export based on current filters.");
                    HideStopButton();
                    return;
                }

                AppendResults($"Found {buildingBlocks.Count} Building Blocks to export based on current filters");
                AppendResults("");

                // Export each Building Block
                for (int i = 0; i < buildingBlocks.Count; i++)
                {
                    // Check for cancellation
                    if (cancellationTokenSource?.Token.IsCancellationRequested == true)
                    {
                        AppendResults("Export operation cancelled by user.");
                        break;
                    }

                    var bb = buildingBlocks[i];
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / buildingBlocks.Count * 100);
                        UpdateStatus($"Exporting: {i + 1} of {buildingBlocks.Count}");
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

                        // Check for file conflicts and stash them
                        if (File.Exists(outputFilePath))
                        {
                            conflictedFiles.Add(outputFilePath);
                            AppendResults($"  ⚠ File already exists, skipping: {Path.GetFileName(outputFilePath)}");
                            continue;
                        }

                        // Export the Building Block
                        wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputFilePath);
                        
                        successCount++;
                        var displayPath = GetRelativePath(exportPath, outputFilePath);
                        logger.LogExport(bb.Name, bb.Category, displayPath);
                        AppendResults($"  ✓ Exported to {displayPath}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        logger.LogError("Export", bb.Name, bb.Category, ex.Message);
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
                HideStopButton();
            }

            var processingTime = (DateTime.Now - startTime).TotalSeconds;
            
            AppendResults("");
            AppendResults("Export Operation Completed");
            AppendResults($"Building Blocks Successfully Exported: {successCount}");
            AppendResults($"Files with Errors: {errorCount}");
            AppendResults($"Logs saved to: {logger.GetLogDirectory()}");
            AppendResults($"Directories Created: {directoriesCreated}");
            AppendResults($"Processing Time: {processingTime:F1} seconds");
            
            // Handle conflicted files
            HandleExportConflicts();
            
            // Export and error logging handled individually above
            
            progressBar.Value = 0;
            UpdateStatus(errorCount == 0 ? "Export completed successfully" : $"Export completed with {errorCount} errors");
        }

        private void BtnExportSelected_Click(object sender, EventArgs e)
        {
            if (!ValidatePaths()) return;

            // Check if any building blocks are selected in the ListView
            if (listViewTemplate.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select one or more Building Blocks from the Template tab to export.", 
                    "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Get selected building blocks from ListView
            var selectedBlocks = new List<BuildingBlockInfo>();
            foreach (ListViewItem item in listViewTemplate.SelectedItems)
            {
                if (item.Tag is BuildingBlockInfo bb)
                {
                    selectedBlocks.Add(bb);
                }
            }

            if (selectedBlocks.Count == 0)
            {
                MessageBox.Show("No valid Building Blocks selected.", 
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Use export directory if set, otherwise prompt
            string exportPath;
            if (!string.IsNullOrEmpty(fullExportDirectoryPath) && Directory.Exists(fullExportDirectoryPath))
            {
                exportPath = fullExportDirectoryPath;
            }
            else
            {
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
                    dialog.SelectedPath = !string.IsNullOrEmpty(fullExportDirectoryPath) ? fullExportDirectoryPath : fullSourceDirectoryPath;
                    
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    exportPath = dialog.SelectedPath;
                }
            }

            UpdateStatus("Exporting selected Building Blocks...");
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.Value = 0;
            
            // Clear conflict list for this operation
            conflictedFiles.Clear();
            
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
                if (!HandleTemplateFileLock(fullTemplatePath, "Export"))
                    return;

                // Initialize WordManager
                wordManager = new WordManager(fullTemplatePath);

                // Export each selected Building Block
                for (int i = 0; i < selectedBlocks.Count; i++)
                {
                    var bb = selectedBlocks[i];
                    
                    try
                    {
                        progressBar.Value = (int)((double)(i + 1) / selectedBlocks.Count * 100);
                        UpdateStatus($"Exporting: {i + 1} of {selectedBlocks.Count}");
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

                        // Check for file conflicts and stash them
                        if (File.Exists(outputFilePath))
                        {
                            conflictedFiles.Add(outputFilePath);
                            AppendResults($"  ⚠ File already exists, skipping: {Path.GetFileName(outputFilePath)}");
                            continue;
                        }

                        // Export the Building Block
                        wordManager.ExportBuildingBlock(bb.Name, bb.Category, outputFilePath);
                        
                        successCount++;
                        var displayPath = GetRelativePath(exportPath, outputFilePath);
                        logger.LogExport(bb.Name, bb.Category, displayPath);
                        AppendResults($"  ✓ Exported to {displayPath}");
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        logger.LogError("Export", bb.Name, bb.Category, ex.Message);
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
            AppendResults($"Logs saved to: {logger.GetLogDirectory()}");
            AppendResults($"Directories Created: {directoriesCreated}");
            AppendResults($"Processing Time: {processingTime:F1} seconds");
            
            // Handle conflicted files
            HandleExportConflicts();
            
            // Export and error logging handled individually above
            
            progressBar.Value = 0;
            UpdateStatus(errorCount == 0 ? "Export completed successfully" : $"Export completed with {errorCount} errors");
        }

        private void BtnRollback_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(fullTemplatePath))
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
                AppendResults($"Template: {Path.GetFileName(fullTemplatePath)}");
                AppendResults("Searching for backup files...");

                WordManager wordManager = null;
                
                try
                {
                    wordManager = new WordManager(fullTemplatePath);
                    
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

        private void LoadSettings()
        {
            settings = Settings.Load();
            
            UpdateTemplatePathDisplay(settings.LastTemplatePath);
            UpdateSourceDirectoryDisplay(settings.LastSourceDirectory);
            UpdateExportDirectoryDisplay(settings.LastExportDirectory);
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
            
            settings.LastTemplatePath = fullTemplatePath;
            settings.LastSourceDirectory = fullSourceDirectoryPath;
            settings.LastExportDirectory = fullExportDirectoryPath;
            settings.FlatImport = chkFlatImport.Checked;
            settings.FlatExport = chkFlatExport.Checked;
            
            settings.Save();
        }

        protected override void OnFormClosing(System.Windows.Forms.FormClosingEventArgs e)
        {
            SaveSettings();
            logger?.Info("Building Blocks Manager closing");
            logger?.EndSession();
            base.OnFormClosing(e);
        }

        private List<FileManager.FileInfo> GetCheckedFiles(TreeNodeCollection nodes)
        {
            var checkedFiles = new List<FileManager.FileInfo>();
            
            foreach (TreeNode node in nodes)
            {
                // If this node represents a file and is checked
                if (node.Checked && node.Tag is FileManager.FileInfo fileInfo)
                {
                    checkedFiles.Add(fileInfo);
                }
                
                // Recursively check child nodes
                if (node.Nodes.Count > 0)
                {
                    checkedFiles.AddRange(GetCheckedFiles(node.Nodes));
                }
            }
            
            return checkedFiles;
        }

        private string ConvertCategoryToPath(string category)
        {
            // Convert category back to folder path structure
            if (string.IsNullOrWhiteSpace(category))
                return ""; // No subfolder for empty categories
                
            // Convert underscores back to spaces in folder names
            return category.Replace('_', ' ');
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

        private void HandleExportConflicts()
        {
            if (conflictedFiles.Count == 0)
                return;

            AppendResults("");
            AppendResults($"⚠ FILE CONFLICTS DETECTED: {conflictedFiles.Count} files already exist in the export directory");
            AppendResults("These files were NOT exported to avoid overwriting existing data:");
            
            foreach (var file in conflictedFiles)
            {
                AppendResults($"  • {Path.GetFileName(file)}");
            }
            
            AppendResults("");
            
            var result = MessageBox.Show(
                $"{conflictedFiles.Count} files already exist in the export directory and were skipped.\n\n" +
                "Do you want to overwrite these existing files?\n\n" +
                "• YES: Overwrite all conflicted files\n" +
                "• NO: Skip these files (current choice)",
                "File Conflicts - Overwrite Existing Files?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2);
            
            if (result == DialogResult.Yes)
            {
                OverwriteConflictedFiles();
            }
            else
            {
                AppendResults("User chose to skip conflicted files. No files were overwritten.");
            }
        }

        private void OverwriteConflictedFiles()
        {
            AppendResults("User chose to overwrite existing files. Re-exporting conflicted items...");
            
            WordManager wordManager = null;
            int overwriteSuccessCount = 0;
            int overwriteErrorCount = 0;
            
            try
            {
                wordManager = new WordManager(fullTemplatePath);
                
                foreach (var conflictedFile in conflictedFiles)
                {
                    try
                    {
                        // Find the building block that corresponds to this file
                        var fileName = Path.GetFileNameWithoutExtension(conflictedFile);
                        if (fileName.StartsWith("AT_"))
                        {
                            var blockName = fileName.Substring(3); // Remove "AT_" prefix
                            
                            // Find the building block in the filtered list
                            var buildingBlock = GetFilteredBuildingBlocks()
                                .FirstOrDefault(bb => bb.Name == blockName);
                            
                            if (buildingBlock != null)
                            {
                                // Delete the existing file
                                File.Delete(conflictedFile);
                                
                                // Export the building block
                                wordManager.ExportBuildingBlock(buildingBlock.Name, buildingBlock.Category, conflictedFile);
                                
                                var displayPath = GetRelativePath(Path.GetDirectoryName(conflictedFile), conflictedFile);
                                logger.LogExport(buildingBlock.Name, buildingBlock.Category, displayPath);
                                AppendResults($"  ✓ Overwritten: {Path.GetFileName(conflictedFile)}");
                                overwriteSuccessCount++;
                            }
                            else
                            {
                                AppendResults($"  ⚠ Could not find building block for: {Path.GetFileName(conflictedFile)}");
                                overwriteErrorCount++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendResults($"  ✗ Failed to overwrite {Path.GetFileName(conflictedFile)}: {ex.Message}");
                        overwriteErrorCount++;
                    }
                }
                
                AppendResults("");
                AppendResults($"Overwrite Summary: Success: {overwriteSuccessCount}, Errors: {overwriteErrorCount}");
            }
            catch (Exception ex)
            {
                AppendResults($"Error during overwrite operation: {ex.Message}");
            }
            finally
            {
                wordManager?.Dispose();
            }
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
                var logDirectory = logger.GetLogDirectory();
                
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

        private void LoggingConfigMenuItem_Click(object sender, EventArgs e)
        {
            var configForm = new LoggingConfigForm(settings);
            if (configForm.ShowDialog() == DialogResult.OK)
            {
                // Settings were changed, reinitialize logger
                InitializeLogger();
                AppendResults("Logging configuration updated. New settings will take effect for the next session.");
            }
        }

        private void ImportRulesMenuItem_Click(object sender, EventArgs e)
        {
            var helpMessage = @"IMPORT/EXPORT RULES

IMPORT (FILE -> TEMPLATE):
• Files must start with 'AT_' to be processed
• Directory structure converts to Building Block categories
• Top-level source directory is ignored in category path

IMPORT EXAMPLE:
• File: C:\MyDocs\AutotextRepo\Legal\Contracts\AT_Agreement.docx
• Selected source directory: AutotextRepo
• Imported autotext name: Agreement
• Imported autotext category: Legal\Contracts
• Imported gallery: AutoText (default)
• Stored in: Selected template file

SPECIAL CHARACTERS:
• Spaces in folder names become underscores in categories
• Invalid filename characters are flagged but processing continues
• Case is preserved in both folder names and Building Block names

FLAT STRUCTURE OPTIONS:
• Flat Import: All files go into single specified category
• Flat Export: All Building Blocks export to single folder (no subfolders)

BACKUP PROCESS:
• The template file is backed up before importing
• Location: Same folder as your template file
• Recovery: Use File → Rollback to restore from most recent backup
• Cleanup: Last 5 backups kept, older ones deleted automatically";

            MessageBox.Show(helpMessage, "Import/Export Rules", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnQueryTemplate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(fullTemplatePath))
            {
                MessageBox.Show("Please select a template file first.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(fullTemplatePath))
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
                if (!HandleTemplateFileLock(fullTemplatePath, "Query Template"))
                {
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.Value = 0;
                    UpdateStatus("Query cancelled");
                    return;
                }

                using (var wordManager = new WordManager(fullTemplatePath))
                {
                    var buildingBlocks = wordManager.GetBuildingBlocks();
                    
                    // Store all building blocks for filtering
                    allBuildingBlocks = buildingBlocks;
                    
                    // Reset filters when loading new template and apply defaults
                    selectedCategories.Clear();
                    selectedGalleries.Clear();
                    selectedTemplates.Clear();
                    
                    // Categories: All checked EXCEPT "System/Hex Entries" (if it exists)
                    var categories = GetUniqueCategories();
                    foreach (var category in categories)
                    {
                        if (category != "System/Hex Entries")
                        {
                            selectedCategories.Add(category);
                        }
                    }
                    
                    // Galleries: All checked EXCEPT "Placeholder" (if it exists)
                    var galleries = GetUniqueGalleries();
                    foreach (var gallery in galleries)
                    {
                        if (gallery != "Placeholder")
                        {
                            selectedGalleries.Add(gallery);
                        }
                    }
                    
                    // Templates: All checked
                    var templates = GetUniqueTemplates();
                    selectedTemplates.AddRange(templates);
                    
                    UpdateFilterButtonText();
                    
                    if (buildingBlocks.Count == 0)
                    {
                        // Clear previous data
                        listViewTemplate.Items.Clear();
                        lblTemplateCount.Text = "No Building Blocks found";
                        
                        UpdateStatus("No Building Blocks found in template");
                        tabControl.SelectedTab = tabTemplate;
                        MessageBox.Show("No Building Blocks found in the template.\n\nThis could mean:\n• The template has no Building Blocks\n• The Building Blocks are in a different format", 
                            "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // Apply current filter (system/hex entries excluded by default)
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
            
            if (!Directory.Exists(fullSourceDirectoryPath))
                return;
                
            // Use standard DirectoryInfo approach
            var rootDir = new DirectoryInfo(fullSourceDirectoryPath);
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
            var templates = GetUniqueTemplates();
            
            var form = new Form
            {
                Text = "Filter Building Blocks",
                Size = new System.Drawing.Size(380, 400),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            // Categories section (left side)
            var lblCategories = new Label
            {
                Text = "Categories:",
                Location = new System.Drawing.Point(15, 10),
                Size = new System.Drawing.Size(70, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            var listBoxCategories = new CheckedListBox
            {
                Location = new System.Drawing.Point(15, 30),
                Size = new System.Drawing.Size(150, 250),
                CheckOnClick = true
            };

            foreach (var category in categories)
            {
                bool isChecked = selectedCategories.Contains(category);
                listBoxCategories.Items.Add(category, isChecked);
            }

            // Galleries section (top right)
            var lblGalleries = new Label
            {
                Text = "Galleries:",
                Location = new System.Drawing.Point(175, 10),
                Size = new System.Drawing.Size(60, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            var listBoxGalleries = new CheckedListBox
            {
                Location = new System.Drawing.Point(175, 30),
                Size = new System.Drawing.Size(180, 100),
                CheckOnClick = true
            };

            foreach (var gallery in galleries)
            {
                bool isChecked = selectedGalleries.Contains(gallery);
                listBoxGalleries.Items.Add(gallery, isChecked);
            }

            // Templates section (bottom right)
            var lblTemplates = new Label
            {
                Text = "Templates:",
                Location = new System.Drawing.Point(175, 160),
                Size = new System.Drawing.Size(70, 20),
                Font = new System.Drawing.Font(Label.DefaultFont, System.Drawing.FontStyle.Bold)
            };

            var listBoxTemplates = new CheckedListBox
            {
                Location = new System.Drawing.Point(175, 180),
                Size = new System.Drawing.Size(180, 80),
                CheckOnClick = true
            };

            foreach (var template in templates)
            {
                bool isChecked = selectedTemplates.Contains(template);
                listBoxTemplates.Items.Add(template, isChecked);
            }

            // Buttons
            var btnSelectAllCat = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(15, 285),
                Size = new System.Drawing.Size(45, 25)
            };

            var btnSelectNoneCat = new Button
            {
                Text = "None",
                Location = new System.Drawing.Point(70, 285),
                Size = new System.Drawing.Size(45, 25)
            };

            var btnSelectAllGal = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(175, 135),
                Size = new System.Drawing.Size(45, 25)
            };

            var btnSelectNoneGal = new Button
            {
                Text = "None",
                Location = new System.Drawing.Point(230, 135),
                Size = new System.Drawing.Size(45, 25)
            };

            var btnSelectAllTmp = new Button
            {
                Text = "All",
                Location = new System.Drawing.Point(175, 265),
                Size = new System.Drawing.Size(45, 25)
            };

            var btnSelectNoneTmp = new Button
            {
                Text = "None",
                Location = new System.Drawing.Point(230, 265),
                Size = new System.Drawing.Size(45, 25)
            };

            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(195, 325),
                Size = new System.Drawing.Size(70, 25),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(275, 325),
                Size = new System.Drawing.Size(70, 25),
                DialogResult = DialogResult.Cancel
            };

            // Event handlers
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

            btnSelectAllTmp.Click += (s, e) => {
                for (int i = 0; i < listBoxTemplates.Items.Count; i++)
                    listBoxTemplates.SetItemChecked(i, true);
            };

            btnSelectNoneTmp.Click += (s, e) => {
                for (int i = 0; i < listBoxTemplates.Items.Count; i++)
                    listBoxTemplates.SetItemChecked(i, false);
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
                selectedTemplates.Clear();
                foreach (string item in listBoxTemplates.CheckedItems)
                {
                    selectedTemplates.Add(item);
                }
            };

            form.Controls.AddRange(new Control[] { 
                lblCategories, listBoxCategories, lblGalleries, listBoxGalleries, lblTemplates, listBoxTemplates,
                btnSelectAllCat, btnSelectNoneCat, btnSelectAllGal, btnSelectNoneGal, btnSelectAllTmp, btnSelectNoneTmp,
                btnOK, btnCancel });
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

        private List<string> GetUniqueTemplates()
        {
            var templates = allBuildingBlocks
                .Select(bb => string.IsNullOrEmpty(bb.Template) ? "(No Template)" : bb.Template)
                .Distinct()
                .OrderBy(tmp => tmp)
                .ToList();

            return templates;
        }

        private bool IsSystemEntry(BuildingBlockInfo bb)
        {
            // More targeted system entry detection
            return bb.Name.Length > 15 && bb.Name.All(c => "0123456789ABCDEF-".Contains(char.ToUpper(c))) ||
                   bb.Name.StartsWith("_") ||
                   (bb.Category != null && bb.Category.Contains("System"));
        }

        private List<BuildingBlockInfo> GetFilteredBuildingBlocks()
        {
            return allBuildingBlocks
                .Where(bb => selectedCategories.Contains(IsSystemEntry(bb) ? "System/Hex Entries" : 
                                                        (string.IsNullOrEmpty(bb.Category) ? "(No Category)" : bb.Category)))
                .Where(bb => selectedGalleries.Contains(string.IsNullOrEmpty(bb.Gallery) ? "(No Gallery)" : bb.Gallery))
                .Where(bb => selectedTemplates.Contains(string.IsNullOrEmpty(bb.Template) ? "(No Template)" : bb.Template))
                .ToList();
        }

        private void ApplyTemplateFilter()
        {
            listViewTemplate.Items.Clear();
            
            var filteredBlocks = GetFilteredBuildingBlocks();

            // Populate ListView with filtered building blocks
            foreach (var bb in filteredBlocks.OrderBy(bb => bb.Name))
            {
                var item = new ListViewItem(bb.Name);
                item.SubItems.Add(bb.Gallery);
                item.SubItems.Add(bb.Category);
                item.SubItems.Add(bb.Template);
                item.Tag = bb;
                listViewTemplate.Items.Add(item);
            }

            UpdateTemplateCount(filteredBlocks.Count);
        }

        private void UpdateFilterButtonText()
        {
            int totalFilters = selectedCategories.Count + selectedGalleries.Count + selectedTemplates.Count;
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


        private ContextMenuStrip CreateTemplateContextMenu()
        {
            var contextMenu = new ContextMenuStrip();
            
            var deleteMenuItem = new ToolStripMenuItem("Delete Building Block")
            {
                ShortcutKeys = Keys.Delete
            };
            deleteMenuItem.Click += (sender, e) => DeleteSelectedTemplate();
            
            contextMenu.Items.Add(deleteMenuItem);
            
            return contextMenu;
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listViewTemplate.Items)
            {
                item.Selected = true;
            }
        }

        private void ListViewTemplate_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                DeleteSelectedTemplate();
                e.Handled = true;
            }
        }


        private void DeleteSelectedTemplate()
        {
            if (listViewTemplate.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select Building Block(s) to delete.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Collect selected building blocks
            var selectedBlocks = new System.Collections.Generic.List<BuildingBlockInfo>();
            foreach (ListViewItem item in listViewTemplate.SelectedItems)
            {
                var buildingBlock = item.Tag as BuildingBlockInfo;
                if (buildingBlock != null)
                {
                    selectedBlocks.Add(buildingBlock);
                }
            }

            if (selectedBlocks.Count == 0)
            {
                MessageBox.Show("Invalid Building Block selection.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Confirm deletion
            string message = selectedBlocks.Count == 1
                ? $"Are you sure you want to delete the Building Block '{selectedBlocks[0].Name}' from category '{selectedBlocks[0].Category}'?\n\nThis action cannot be undone."
                : $"Are you sure you want to delete {selectedBlocks.Count} Building Blocks?\n\nThis action cannot be undone.";

            var result = MessageBox.Show(message, "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
                return;

            int successCount = 0;
            int errorCount = 0;
            var itemsToRemove = new System.Collections.Generic.List<ListViewItem>();

            try
            {
                using (var wordManager = new WordManager(fullTemplatePath))
                {
                    // Create backup before deleting
                    wordManager.CreateBackup();
                    
                    // Find the ListView items corresponding to selected building blocks
                    foreach (var buildingBlock in selectedBlocks)
                    {
                        var correspondingItem = listViewTemplate.Items.Cast<ListViewItem>()
                            .FirstOrDefault(item => item.Tag is BuildingBlockInfo bb && 
                                           bb.Name == buildingBlock.Name && bb.Category == buildingBlock.Category);
                        
                        try
                        {
                            wordManager.DeleteBuildingBlock(buildingBlock.Name, buildingBlock.Category);
                            AppendResults($"Deleted Building Block: {buildingBlock.Name} from category {buildingBlock.Category}");
                            logger.LogDeletion(buildingBlock.Name, buildingBlock.Category);
                            
                            // Remove from the data collections
                            allBuildingBlocks.RemoveAll(bb => bb.Name == buildingBlock.Name && bb.Category == buildingBlock.Category);
                            
                            // Mark item for removal from ListView
                            if (correspondingItem != null)
                            {
                                itemsToRemove.Add(correspondingItem);
                            }
                            
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            AppendResults($"Failed to delete {buildingBlock.Name}: {ex.Message}");
                            logger.LogError("Delete", buildingBlock.Name, buildingBlock.Category, ex.Message);
                            errorCount++;
                        }
                    }
                    
                    // Remove successfully deleted items from ListView immediately
                    foreach (var item in itemsToRemove)
                    {
                        listViewTemplate.Items.Remove(item);
                    }
                    
                    // Update template count display
                    UpdateTemplateCount(listViewTemplate.Items.Count);
                    
                    UpdateStatus($"Deleted {successCount} Building Block(s)" + (errorCount > 0 ? $" ({errorCount} errors)" : ""));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to delete Building Block(s): {ex.Message}", "Delete Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendResults($"Failed to delete Building Blocks: {ex.Message}");
            }
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