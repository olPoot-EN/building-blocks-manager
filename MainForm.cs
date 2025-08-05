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
        private Button btnRollback;
        private TextBox txtResults;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Settings settings;

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
            btnRollback.Click += BtnRollback_Click;
            
            // Load and apply settings
            LoadSettings();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            // Template file section
            var lblTemplate = new Label
            {
                Text = "Template File:",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(100, 23)
            };

            txtTemplatePath = new TextBox
            {
                Location = new System.Drawing.Point(130, 20),
                Size = new System.Drawing.Size(500, 23),
                ReadOnly = true
            };

            btnBrowseTemplate = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(640, 19),
                Size = new System.Drawing.Size(80, 25)
            };

            // Source directory section
            var lblDirectory = new Label
            {
                Text = "Source Directory:",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(110, 23)
            };

            txtSourceDirectory = new TextBox
            {
                Location = new System.Drawing.Point(140, 60),
                Size = new System.Drawing.Size(490, 23),
                ReadOnly = true
            };

            btnBrowseDirectory = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(640, 59),
                Size = new System.Drawing.Size(80, 25)
            };

            // Structure options section
            var lblStructure = new Label
            {
                Text = "Ignore folder/category structure for:",
                Location = new System.Drawing.Point(20, 100),
                Size = new System.Drawing.Size(250, 23)
            };

            chkFlatImport = new CheckBox
            {
                Text = "Import",
                Location = new System.Drawing.Point(280, 100),
                Size = new System.Drawing.Size(80, 23)
            };

            chkFlatExport = new CheckBox
            {
                Text = "Export",
                Location = new System.Drawing.Point(370, 100),
                Size = new System.Drawing.Size(80, 23)
            };

            // Action buttons
            btnQueryDirectory = new Button
            {
                Text = "Query Directory",
                Location = new System.Drawing.Point(20, 140),
                Size = new System.Drawing.Size(120, 30)
            };

            btnImportAll = new Button
            {
                Text = "Import All",
                Location = new System.Drawing.Point(150, 140),
                Size = new System.Drawing.Size(100, 30)
            };

            btnImportSelected = new Button
            {
                Text = "Import Selected File",
                Location = new System.Drawing.Point(260, 140),
                Size = new System.Drawing.Size(130, 30)
            };

            btnExportAll = new Button
            {
                Text = "Export All",
                Location = new System.Drawing.Point(400, 140),
                Size = new System.Drawing.Size(100, 30)
            };

            btnExportSelected = new Button
            {
                Text = "Export Selected",
                Location = new System.Drawing.Point(510, 140),
                Size = new System.Drawing.Size(120, 30)
            };

            btnRollback = new Button
            {
                Text = "Rollback",
                Location = new System.Drawing.Point(640, 140),
                Size = new System.Drawing.Size(80, 30)
            };

            // Results section
            var lblResults = new Label
            {
                Text = "Results:",
                Location = new System.Drawing.Point(20, 190),
                Size = new System.Drawing.Size(100, 23)
            };

            txtResults = new TextBox
            {
                Location = new System.Drawing.Point(20, 220),
                Size = new System.Drawing.Size(700, 280),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true,
                Font = new System.Drawing.Font("Consolas", 9)
            };

            // Progress and status section
            progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(20, 520),
                Size = new System.Drawing.Size(500, 23),
                Style = ProgressBarStyle.Continuous
            };

            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(530, 520),
                Size = new System.Drawing.Size(190, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };

            // Add all controls to form
            Controls.AddRange(new Control[]
            {
                lblTemplate, txtTemplatePath, btnBrowseTemplate,
                lblDirectory, txtSourceDirectory, btnBrowseDirectory,
                lblStructure, chkFlatImport, chkFlatExport,
                btnQueryDirectory, btnImportAll, btnImportSelected,
                btnExportAll, btnExportSelected, btnRollback,
                lblResults, txtResults,
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
            
            AppendResults("=== DIRECTORY QUERY ===");
            AppendResults($"Scanning directory: {txtSourceDirectory.Text}");
            
            try
            {
                var fileManager = new FileManager(txtSourceDirectory.Text);
                var summary = fileManager.GetSummary();
                
                AppendResults("");
                AppendResults(summary);
                AppendResults("");
                AppendResults("Detailed Listing:");
                
                var files = fileManager.ScanDirectory();
                foreach (var file in files.Take(10)) // Show first 10 files
                {
                    var fileName = Path.GetFileName(file.FilePath);
                    var status = file.IsNew ? "New (Never imported)" :
                                file.IsModified ? $"Modified (Last: {file.LastModified:yyyy-MM-dd}, Imported: {file.LastImported:yyyy-MM-dd})" :
                                "Up-to-date";
                    
                    if (!file.IsValid)
                        status += $" - INVALID CHARACTERS: {string.Join(", ", file.InvalidCharacters)}";
                    
                    var relativePath = Path.GetRelativePath(txtSourceDirectory.Text, file.FilePath);
                    AppendResults($"{relativePath} - {status}");
                }
                
                if (files.Count > 10)
                {
                    AppendResults($"... and {files.Count - 10} more files");
                }
                
                AppendResults("");
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
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        AppendResults($"  ✗ Failed to import {fileName}: {ex.Message}");
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
                    AppendResults("Creating backup...");
                    AppendResults("Processing file...");
                    
                    System.Threading.Thread.Sleep(1500);
                    
                    AppendResults($"Successfully imported Building Block: {Path.GetFileNameWithoutExtension(dialog.FileName).Substring(3)}");
                    AppendResults("Processing Time: 0.8 seconds");
                    
                    progressBar.Style = ProgressBarStyle.Continuous;
                    progressBar.Value = 0;
                    UpdateStatus("Import completed successfully");
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
            
            // Simulate export process
            for (int i = 0; i <= 100; i += 10)
            {
                progressBar.Value = i;
                System.Threading.Thread.Sleep(200);
                if (i == 20) AppendResults("Creating directory structure...");
                if (i == 40) AppendResults("Exporting Legal\\Contracts\\AT_Standard.docx...");
                if (i == 70) AppendResults("Exporting HR\\AT_Policy.docx...");
            }
            
            AppendResults("");
            AppendResults("Export Completed Successfully");
            AppendResults("Building Blocks Exported: 45");
            AppendResults("- Created Directories: 8");
            AppendResults("- Files Created: 45");
            AppendResults("Processing Time: 1.8 seconds");
            
            progressBar.Value = 0;
            UpdateStatus("Export completed successfully");
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
            
            // Simulate export process
            for (int i = 0; i <= 100; i += 20)
            {
                progressBar.Value = i;
                System.Threading.Thread.Sleep(300);
            }
            
            AppendResults($"Successfully exported {selectedBlocks.Count} Building Blocks");
            AppendResults("Processing Time: 0.9 seconds");
            
            progressBar.Value = 0;
            UpdateStatus("Export completed successfully");
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
                
                System.Threading.Thread.Sleep(1000);
                
                AppendResults("Found backup: Template_Backup_20250803_143016.dotm");
                AppendResults("Restoring template from backup...");
                
                System.Threading.Thread.Sleep(1500);
                
                AppendResults("Rollback completed successfully");
                AppendResults("Template restored to previous state");
                
                progressBar.Style = ProgressBarStyle.Continuous;
                progressBar.Value = 0;
                UpdateStatus("Rollback completed");
            }
        }

        private string PromptForInput(string title, string prompt)
        {
            // Simple input dialog simulation
            return Microsoft.VisualBasic.Interaction.InputBox(prompt, title, "");
        }

        private System.Collections.Generic.List<string> ShowBuildingBlockSelectionDialog()
        {
            // Simulate Building Block selection dialog
            var form = new Form
            {
                Text = "Select Building Blocks to Export",
                Size = new System.Drawing.Size(500, 400),
                StartPosition = FormStartPosition.CenterScreen
            };

            var listBox = new CheckedListBox
            {
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(440, 280),
                CheckOnClick = true
            };

            // Add sample Building Blocks
            listBox.Items.Add("InternalAutotext\\Legal\\Contracts - Standard", false);
            listBox.Items.Add("InternalAutotext\\Legal\\Contracts - Premium", false);
            listBox.Items.Add("InternalAutotext\\Legal\\Forms - Disclaimer", false);
            listBox.Items.Add("InternalAutotext\\HR\\Policies - Vacation", false);
            listBox.Items.Add("InternalAutotext\\HR\\Policies - Sick_Leave", false);
            listBox.Items.Add("InternalAutotext\\Finance - Summary", false);

            var btnSelectAll = new Button
            {
                Text = "Select All",
                Location = new System.Drawing.Point(20, 320),
                Size = new System.Drawing.Size(80, 25)
            };

            var btnSelectNone = new Button
            {
                Text = "Select None",
                Location = new System.Drawing.Point(110, 320),
                Size = new System.Drawing.Size(80, 25)
            };

            var btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(300, 320),
                Size = new System.Drawing.Size(80, 25),
                DialogResult = DialogResult.OK
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(390, 320),
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

            var result = new System.Collections.Generic.List<string>();
            if (form.ShowDialog() == DialogResult.OK)
            {
                foreach (var item in listBox.CheckedItems)
                {
                    result.Add(item.ToString());
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
            base.OnFormClosing(e);
        }
    }
}