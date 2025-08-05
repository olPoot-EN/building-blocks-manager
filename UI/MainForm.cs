using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using BuildingBlocksManager.Core;

namespace BuildingBlocksManager.UI
{
    public partial class MainForm : Form
    {
        private FileManager _fileManager;
        private WordManager _wordManager;
        private ImportTracker _importTracker;
        private Logger _logger;
        private ImportExportManager _importExportManager;
        private SettingsManager _settingsManager;

        // UI Controls
        private TextBox _sourceDirectoryTextBox;
        private TextBox _templateFileTextBox;
        private CheckBox _flatImportCheckBox;
        private CheckBox _flatExportCheckBox;
        private TextBox _resultsTextBox;
        private ProgressBar _progressBar;
        private Label _statusLabel;
        private Button _queryButton;
        private Button _importAllButton;
        private Button _importSelectedButton;
        private Button _exportAllButton;
        private Button _exportSelectedButton;
        private Button _rollbackButton;

        public MainForm()
        {
            InitializeComponents();
            InitializeCore();
            LoadSettings();
            SetupEventHandlers();
        }

        private void InitializeComponents()
        {
            Text = "Building Blocks Manager";
            Size = new Size(800, 600);
            MinimumSize = new Size(600, 400);
            StartPosition = FormStartPosition.CenterScreen;

            var mainPanel = new TableLayoutPanel();
            mainPanel.Dock = DockStyle.Fill;
            mainPanel.ColumnCount = 1;
            mainPanel.RowCount = 4;
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // File Selection Panel
            var fileSelectionPanel = CreateFileSelectionPanel();
            mainPanel.Controls.Add(fileSelectionPanel, 0, 0);

            // Action Buttons Panel
            var actionPanel = CreateActionButtonsPanel();
            mainPanel.Controls.Add(actionPanel, 0, 1);

            // Results Panel
            var resultsPanel = CreateResultsPanel();
            mainPanel.Controls.Add(resultsPanel, 0, 2);

            // Status Panel
            var statusPanel = CreateStatusPanel();
            mainPanel.Controls.Add(statusPanel, 0, 3);

            Controls.Add(mainPanel);

            // Menu Bar
            var menuStrip = CreateMenuStrip();
            MainMenuStrip = menuStrip;
            Controls.Add(menuStrip);
        }

        private Panel CreateFileSelectionPanel()
        {
            var panel = new Panel();
            panel.Height = 120;
            panel.Dock = DockStyle.Top;
            panel.Padding = new Padding(10);

            // Source Directory
            var sourceLabel = new Label();
            sourceLabel.Text = "Source Directory:";
            sourceLabel.Location = new Point(10, 15);
            sourceLabel.Size = new Size(100, 20);

            _sourceDirectoryTextBox = new TextBox();
            _sourceDirectoryTextBox.Location = new Point(120, 12);
            _sourceDirectoryTextBox.Size = new Size(500, 20);
            _sourceDirectoryTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            var sourceBrowseButton = new Button();
            sourceBrowseButton.Text = "Browse";
            sourceBrowseButton.Location = new Point(630, 10);
            sourceBrowseButton.Size = new Size(75, 25);
            sourceBrowseButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            sourceBrowseButton.Click += SourceBrowseButton_Click;

            // Template File
            var templateLabel = new Label();
            templateLabel.Text = "Template File:";
            templateLabel.Location = new Point(10, 45);
            templateLabel.Size = new Size(100, 20);

            _templateFileTextBox = new TextBox();
            _templateFileTextBox.Location = new Point(120, 42);
            _templateFileTextBox.Size = new Size(500, 20);
            _templateFileTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            var templateBrowseButton = new Button();
            templateBrowseButton.Text = "Browse";
            templateBrowseButton.Location = new Point(630, 40);
            templateBrowseButton.Size = new Size(75, 25);
            templateBrowseButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            templateBrowseButton.Click += TemplateBrowseButton_Click;

            // Options
            var optionsLabel = new Label();
            optionsLabel.Text = "Options:";
            optionsLabel.Location = new Point(10, 75);
            optionsLabel.Size = new Size(100, 20);

            _flatImportCheckBox = new CheckBox();
            _flatImportCheckBox.Text = "Flat Import (ignore folder structure)";
            _flatImportCheckBox.Location = new Point(120, 75);
            _flatImportCheckBox.Size = new Size(200, 20);

            _flatExportCheckBox = new CheckBox();
            _flatExportCheckBox.Text = "Flat Export (single folder)";
            _flatExportCheckBox.Location = new Point(330, 75);
            _flatExportCheckBox.Size = new Size(200, 20);

            panel.Controls.AddRange(new Control[] {
                sourceLabel, _sourceDirectoryTextBox, sourceBrowseButton,
                templateLabel, _templateFileTextBox, templateBrowseButton,
                optionsLabel, _flatImportCheckBox, _flatExportCheckBox
            });

            return panel;
        }

        private Panel CreateActionButtonsPanel()
        {
            var panel = new Panel();
            panel.Height = 60;
            panel.Dock = DockStyle.Top;
            panel.Padding = new Padding(10);

            _queryButton = new Button();
            _queryButton.Text = "Query Directory";
            _queryButton.Location = new Point(10, 15);
            _queryButton.Size = new Size(120, 30);
            _queryButton.Click += QueryButton_Click;

            _importAllButton = new Button();
            _importAllButton.Text = "Import All";
            _importAllButton.Location = new Point(140, 15);
            _importAllButton.Size = new Size(120, 30);
            _importAllButton.Click += ImportAllButton_Click;

            _importSelectedButton = new Button();
            _importSelectedButton.Text = "Import Selected";
            _importSelectedButton.Location = new Point(270, 15);
            _importSelectedButton.Size = new Size(120, 30);
            _importSelectedButton.Click += ImportSelectedButton_Click;

            _exportAllButton = new Button();
            _exportAllButton.Text = "Export All";
            _exportAllButton.Location = new Point(400, 15);
            _exportAllButton.Size = new Size(120, 30);
            _exportAllButton.Click += ExportAllButton_Click;

            _exportSelectedButton = new Button();
            _exportSelectedButton.Text = "Export Selected";
            _exportSelectedButton.Location = new Point(530, 15);
            _exportSelectedButton.Size = new Size(120, 30);
            _exportSelectedButton.Click += ExportSelectedButton_Click;

            _rollbackButton = new Button();
            _rollbackButton.Text = "Rollback";
            _rollbackButton.Location = new Point(660, 15);
            _rollbackButton.Size = new Size(80, 30);
            _rollbackButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            _rollbackButton.Click += RollbackButton_Click;

            panel.Controls.AddRange(new Control[] {
                _queryButton, _importAllButton, _importSelectedButton,
                _exportAllButton, _exportSelectedButton, _rollbackButton
            });

            return panel;
        }

        private Panel CreateResultsPanel()
        {
            var panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.Padding = new Padding(10);

            var resultsLabel = new Label();
            resultsLabel.Text = "Results:";
            resultsLabel.Location = new Point(10, 10);
            resultsLabel.Size = new Size(100, 20);

            _resultsTextBox = new TextBox();
            _resultsTextBox.Multiline = true;
            _resultsTextBox.ScrollBars = ScrollBars.Vertical;
            _resultsTextBox.ReadOnly = true;
            _resultsTextBox.Location = new Point(10, 35);
            _resultsTextBox.Size = new Size(panel.Width - 20, panel.Height - 50);
            _resultsTextBox.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            _resultsTextBox.Font = new Font("Consolas", 9, FontStyle.Regular);

            panel.Controls.AddRange(new Control[] { resultsLabel, _resultsTextBox });
            return panel;
        }

        private Panel CreateStatusPanel()
        {
            var panel = new Panel();
            panel.Height = 50;
            panel.Dock = DockStyle.Bottom;
            panel.Padding = new Padding(10);

            _progressBar = new ProgressBar();
            _progressBar.Location = new Point(10, 10);
            _progressBar.Size = new Size(400, 20);
            _progressBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            _statusLabel = new Label();
            _statusLabel.Text = "Ready";
            _statusLabel.Location = new Point(10, 35);
            _statusLabel.Size = new Size(400, 15);
            _statusLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            panel.Controls.AddRange(new Control[] { _progressBar, _statusLabel });
            return panel;
        }

        private MenuStrip CreateMenuStrip()
        {
            var menuStrip = new MenuStrip();

            // File Menu
            var fileMenu = new ToolStripMenuItem("File");
            
            var settingsMenuItem = new ToolStripMenuItem("Settings");
            settingsMenuItem.Click += SettingsMenuItem_Click;
            
            var viewLogMenuItem = new ToolStripMenuItem("View Log File");
            viewLogMenuItem.Click += ViewLogMenuItem_Click;
            
            var exitMenuItem = new ToolStripMenuItem("Exit");
            exitMenuItem.Click += (s, e) => Close();

            fileMenu.DropDownItems.AddRange(new ToolStripItem[] {
                settingsMenuItem, new ToolStripSeparator(), viewLogMenuItem, 
                new ToolStripSeparator(), exitMenuItem
            });

            // Help Menu
            var helpMenu = new ToolStripMenuItem("Help");
            
            var aboutMenuItem = new ToolStripMenuItem("About");
            aboutMenuItem.Click += AboutMenuItem_Click;
            
            helpMenu.DropDownItems.Add(aboutMenuItem);

            menuStrip.Items.AddRange(new ToolStripItem[] { fileMenu, helpMenu });
            return menuStrip;
        }

        private void InitializeCore()
        {
            _settingsManager = new SettingsManager();
            _fileManager = new FileManager();
            _wordManager = new WordManager();
            _importTracker = new ImportTracker();
            _logger = new Logger();
            _importExportManager = new ImportExportManager(_fileManager, _wordManager, _importTracker, _logger);
        }

        private void SetupEventHandlers()
        {
            _importExportManager.ProgressUpdate += OnProgressUpdate;
            _importExportManager.ProgressPercentageUpdate += OnProgressPercentageUpdate;
            
            FormClosing += MainForm_FormClosing;
            Load += MainForm_Load;
        }

        private void LoadSettings()
        {
            var settings = _settingsManager.GetSettings();
            
            _sourceDirectoryTextBox.Text = settings.LastSourceDirectory;
            _templateFileTextBox.Text = settings.LastTemplateFile;
            _flatImportCheckBox.Checked = settings.FlatImportEnabled;
            _flatExportCheckBox.Checked = settings.FlatExportEnabled;

            // Apply window settings
            if (!string.IsNullOrEmpty(settings.WindowSize))
            {
                var sizeParts = settings.WindowSize.Split(',');
                if (sizeParts.Length == 2 && 
                    int.TryParse(sizeParts[0], out int width) && 
                    int.TryParse(sizeParts[1], out int height))
                {
                    Size = new Size(Math.Max(MinimumSize.Width, width), Math.Max(MinimumSize.Height, height));
                }
            }

            if (!string.IsNullOrEmpty(settings.WindowLocation))
            {
                var locationParts = settings.WindowLocation.Split(',');
                if (locationParts.Length == 2 && 
                    int.TryParse(locationParts[0], out int x) && 
                    int.TryParse(locationParts[1], out int y))
                {
                    StartPosition = FormStartPosition.Manual;
                    Location = new Point(x, y);
                }
            }
        }

        private void SaveSettings()
        {
            _settingsManager.UpdateLastSourceDirectory(_sourceDirectoryTextBox.Text);
            _settingsManager.UpdateLastTemplateFile(_templateFileTextBox.Text);
            _settingsManager.UpdateImportOptions(_flatImportCheckBox.Checked, 
                "InternalAutotext", true, true);
            _settingsManager.UpdateExportOptions(_flatExportCheckBox.Checked);
            _settingsManager.UpdateWindowSettings($"{Size.Width},{Size.Height}", $"{Location.X},{Location.Y}");
        }

        private void AppendResults(string text)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendResults), text);
                return;
            }

            _resultsTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {text}{Environment.NewLine}");
            _resultsTextBox.ScrollToCaret();
        }

        private void UpdateStatus(string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(UpdateStatus), status);
                return;
            }

            _statusLabel.Text = status;
        }

        private void UpdateProgress(int percentage)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<int>(UpdateProgress), percentage);
                return;
            }

            _progressBar.Value = Math.Max(0, Math.Min(100, percentage));
        }

        private void OnProgressUpdate(object sender, string message)
        {
            UpdateStatus(message);
            AppendResults(message);
        }

        private void OnProgressPercentageUpdate(object sender, int percentage)
        {
            UpdateProgress(percentage);
        }

        private void EnableControls(bool enabled)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<bool>(EnableControls), enabled);
                return;
            }

            _queryButton.Enabled = enabled;
            _importAllButton.Enabled = enabled;
            _importSelectedButton.Enabled = enabled;
            _exportAllButton.Enabled = enabled;
            _exportSelectedButton.Enabled = enabled;
            _rollbackButton.Enabled = enabled;
        }

        private bool ValidateInputs()
        {
            if (string.IsNullOrWhiteSpace(_sourceDirectoryTextBox.Text))
            {
                MessageBox.Show("Please select a source directory.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!Directory.Exists(_sourceDirectoryTextBox.Text))
            {
                MessageBox.Show("Source directory does not exist.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(_templateFileTextBox.Text))
            {
                MessageBox.Show("Please select a template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            if (!File.Exists(_templateFileTextBox.Text))
            {
                MessageBox.Show("Template file does not exist.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        // Event Handlers
        private void SourceBrowseButton_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select source directory containing AT_*.docx files";
                folderDialog.SelectedPath = _sourceDirectoryTextBox.Text;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    _sourceDirectoryTextBox.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void TemplateBrowseButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Word Templates (*.dotm)|*.dotm|All Files (*.*)|*.*";
                openFileDialog.Title = "Select Word Template File";
                openFileDialog.InitialDirectory = Path.GetDirectoryName(_templateFileTextBox.Text);

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _templateFileTextBox.Text = openFileDialog.FileName;
                }
            }
        }

        private async void QueryButton_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            EnableControls(false);
            _resultsTextBox.Clear();
            UpdateProgress(0);

            try
            {
                var result = await System.Threading.Tasks.Task.Run(() => 
                    _importExportManager.QueryDirectory(_sourceDirectoryTextBox.Text));

                AppendResults($"Directory Query Results:");
                AppendResults($"  New files: {result.NewFilesCount}");
                AppendResults($"  Modified files: {result.ModifiedFilesCount}");
                AppendResults($"  Up-to-date files: {result.UpToDateFilesCount}");
                AppendResults($"  Missing files: {result.MissingFilesCount}");
                AppendResults($"  Ignored files: {result.IgnoredFilesCount}");
                AppendResults($"  Scan time: {result.ScanTime.TotalSeconds:F2} seconds");

                if (result.NewFiles.Any())
                {
                    AppendResults("\nNew files found:");
                    foreach (var file in result.NewFiles)
                    {
                        AppendResults($"  • {Path.GetFileName(file)}");
                    }
                }

                if (result.MissingFiles.Any())
                {
                    AppendResults("\nMissing files:");
                    foreach (var file in result.MissingFiles)
                    {
                        AppendResults($"  • {Path.GetFileName(file)}");
                    }
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Error: {ex.Message}");
                MessageBox.Show($"Query failed: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                UpdateProgress(0);
                UpdateStatus("Ready");
            }
        }

        private async void ImportAllButton_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs()) return;

            EnableControls(false);
            _resultsTextBox.Clear();
            UpdateProgress(0);

            try
            {
                var options = new ImportExportManager.ImportOptions
                {
                    FlatImport = _flatImportCheckBox.Checked,
                    FlatImportCategory = "InternalAutotext",
                    ImportOnlyChanged = true,
                    ShowWarningsForNewFiles = true
                };

                var result = await System.Threading.Tasks.Task.Run(() => 
                    _importExportManager.BatchImport(_sourceDirectoryTextBox.Text, _templateFileTextBox.Text, options));

                AppendResults($"Import Results:");
                AppendResults($"  Success: {result.Success}");
                AppendResults($"  Imported: {result.ImportedCount}");
                AppendResults($"  Failed: {result.FailedCount}");
                AppendResults($"  Skipped: {result.SkippedCount}");
                AppendResults($"  Processing time: {result.ProcessingTime.TotalSeconds:F2} seconds");

                if (result.FailedImports.Any())
                {
                    AppendResults("\nFailed imports:");
                    foreach (var failure in result.FailedImports)
                    {
                        AppendResults($"  • {failure}");
                    }
                }

                if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
                {
                    MessageBox.Show($"Import failed: {result.ErrorMessage}", "Import Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Error: {ex.Message}");
                MessageBox.Show($"Import failed: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                UpdateProgress(0);
                UpdateStatus("Ready");
            }
        }

        private async void ImportSelectedButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_templateFileTextBox.Text) || !File.Exists(_templateFileTextBox.Text))
            {
                MessageBox.Show("Please select a valid template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            EnableControls(false);
            _resultsTextBox.Clear();
            UpdateProgress(0);

            try
            {
                var options = new ImportExportManager.ImportOptions
                {
                    FlatImport = _flatImportCheckBox.Checked,
                    FlatImportCategory = "InternalAutotext",
                    ShowWarningsForNewFiles = true
                };

                var result = await System.Threading.Tasks.Task.Run(() => 
                    _importExportManager.SelectiveImport(_templateFileTextBox.Text, options));

                AppendResults($"Selective Import Results:");
                AppendResults($"  Success: {result.Success}");
                AppendResults($"  Imported: {result.ImportedCount}");
                AppendResults($"  Failed: {result.FailedCount}");

                if (!string.IsNullOrEmpty(result.ErrorMessage))
                {
                    AppendResults($"  Message: {result.ErrorMessage}");
                }

                if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
                {
                    MessageBox.Show($"Import failed: {result.ErrorMessage}", "Import Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Error: {ex.Message}");
                MessageBox.Show($"Import failed: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                UpdateProgress(0);
                UpdateStatus("Ready");
            }
        }

        private async void ExportAllButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_templateFileTextBox.Text) || !File.Exists(_templateFileTextBox.Text))
            {
                MessageBox.Show("Please select a valid template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string exportDirectory;
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select export directory";
                if (folderDialog.ShowDialog() != DialogResult.OK)
                    return;
                
                exportDirectory = folderDialog.SelectedPath;
            }

            EnableControls(false);
            _resultsTextBox.Clear();
            UpdateProgress(0);

            try
            {
                var options = new ImportExportManager.ExportOptions
                {
                    FlatExport = _flatExportCheckBox.Checked,
                    FlatExportDirectory = _flatExportCheckBox.Checked ? exportDirectory : null,
                    HierarchicalExport = !_flatExportCheckBox.Checked,
                    ExportRootDirectory = !_flatExportCheckBox.Checked ? exportDirectory : null
                };

                var result = await System.Threading.Tasks.Task.Run(() => 
                    _importExportManager.BatchExport(_templateFileTextBox.Text, options));

                AppendResults($"Export Results:");
                AppendResults($"  Success: {result.Success}");
                AppendResults($"  Exported: {result.ExportedCount}");
                AppendResults($"  Failed: {result.FailedCount}");
                AppendResults($"  Export directory: {result.ExportDirectory}");
                AppendResults($"  Processing time: {result.ProcessingTime.TotalSeconds:F2} seconds");

                if (result.FailedExports.Any())
                {
                    AppendResults("\nFailed exports:");
                    foreach (var failure in result.FailedExports)
                    {
                        AppendResults($"  • {failure}");
                    }
                }

                if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
                {
                    MessageBox.Show($"Export failed: {result.ErrorMessage}", "Export Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Error: {ex.Message}");
                MessageBox.Show($"Export failed: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                UpdateProgress(0);
                UpdateStatus("Ready");
            }
        }

        private async void ExportSelectedButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_templateFileTextBox.Text) || !File.Exists(_templateFileTextBox.Text))
            {
                MessageBox.Show("Please select a valid template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string exportDirectory;
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select export directory";
                if (folderDialog.ShowDialog() != DialogResult.OK)
                    return;
                
                exportDirectory = folderDialog.SelectedPath;
            }

            EnableControls(false);
            _resultsTextBox.Clear();
            UpdateProgress(0);

            try
            {
                var options = new ImportExportManager.ExportOptions
                {
                    FlatExport = _flatExportCheckBox.Checked,
                    FlatExportDirectory = _flatExportCheckBox.Checked ? exportDirectory : null,
                    HierarchicalExport = !_flatExportCheckBox.Checked,
                    ExportRootDirectory = !_flatExportCheckBox.Checked ? exportDirectory : null
                };

                var result = await System.Threading.Tasks.Task.Run(() => 
                    _importExportManager.SelectiveExport(_templateFileTextBox.Text, options));

                AppendResults($"Selective Export Results:");
                AppendResults($"  Success: {result.Success}");
                AppendResults($"  Exported: {result.ExportedCount}");
                AppendResults($"  Failed: {result.FailedCount}");
                AppendResults($"  Export directory: {result.ExportDirectory}");

                if (!string.IsNullOrEmpty(result.ErrorMessage))
                {
                    AppendResults($"  Message: {result.ErrorMessage}");
                }

                if (!result.Success && !string.IsNullOrEmpty(result.ErrorMessage))
                {
                    MessageBox.Show($"Export failed: {result.ErrorMessage}", "Export Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Error: {ex.Message}");
                MessageBox.Show($"Export failed: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableControls(true);
                UpdateProgress(0);
                UpdateStatus("Ready");
            }
        }

        private void RollbackButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_templateFileTextBox.Text))
            {
                MessageBox.Show("Please select a template file.", "Validation Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                var backups = _wordManager.GetBackupFiles(_templateFileTextBox.Text);
                
                if (!backups.Any())
                {
                    MessageBox.Show("No backup files found for the selected template.", "No Backups", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var dialog = new Form())
                {
                    dialog.Text = "Select Backup to Restore";
                    dialog.Size = new Size(500, 300);
                    dialog.StartPosition = FormStartPosition.CenterParent;

                    var listBox = new ListBox();
                    listBox.Dock = DockStyle.Fill;
                    
                    foreach (var backup in backups)
                    {
                        var fileName = Path.GetFileName(backup);
                        var fileTime = File.GetCreationTime(backup);
                        listBox.Items.Add($"{fileName} ({fileTime:yyyy-MM-dd HH:mm:ss})");
                    }

                    if (listBox.Items.Count > 0)
                        listBox.SelectedIndex = 0;

                    var buttonPanel = new Panel();
                    buttonPanel.Dock = DockStyle.Bottom;
                    buttonPanel.Height = 50;

                    var restoreButton = new Button();
                    restoreButton.Text = "Restore";
                    restoreButton.Location = new Point(10, 10);
                    restoreButton.DialogResult = DialogResult.OK;

                    var cancelButton = new Button();
                    cancelButton.Text = "Cancel";
                    cancelButton.Location = new Point(100, 10);
                    cancelButton.DialogResult = DialogResult.Cancel;

                    buttonPanel.Controls.AddRange(new Control[] { restoreButton, cancelButton });
                    dialog.Controls.AddRange(new Control[] { listBox, buttonPanel });

                    if (dialog.ShowDialog() == DialogResult.OK && listBox.SelectedIndex >= 0)
                    {
                        var selectedBackup = backups[listBox.SelectedIndex];
                        var confirmResult = MessageBox.Show(
                            $"Are you sure you want to restore from backup?\n\nThis will overwrite:\n{_templateFileTextBox.Text}\n\nWith backup:\n{Path.GetFileName(selectedBackup)}",
                            "Confirm Restore",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning
                        );

                        if (confirmResult == DialogResult.Yes)
                        {
                            _wordManager.RestoreFromBackup(selectedBackup, _templateFileTextBox.Text);
                            AppendResults($"Template restored from backup: {Path.GetFileName(selectedBackup)}");
                            MessageBox.Show("Template successfully restored from backup.", "Restore Complete", 
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppendResults($"Rollback error: {ex.Message}");
                MessageBox.Show($"Rollback failed: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SettingsMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Settings dialog not yet implemented.", "Settings", 
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ViewLogMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                var logFile = _logger.GetCurrentLogFilePath();
                if (File.Exists(logFile))
                {
                    System.Diagnostics.Process.Start("notepad.exe", logFile);
                }
                else
                {
                    MessageBox.Show("Log file not found.", "Log File", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open log file: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AboutMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(
                "Building Blocks Manager v1.0\n\n" +
                "A tool for importing and exporting Word Building Blocks\n" +
                "from directory-based AutoText files.\n\n" +
                "© 2025 Internal Tool",
                "About Building Blocks Manager",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            AppendResults("Building Blocks Manager started");
            AppendResults("Ready to process AutoText files");
            
            // Clean up old log files
            try
            {
                _logger.CleanupOldLogs();
            }
            catch (Exception ex)
            {
                AppendResults($"Warning: Could not cleanup old logs: {ex.Message}");
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                SaveSettings();
                _wordManager?.Dispose();
                _logger?.Dispose();
            }
            catch (Exception ex)
            {
                // Don't prevent closing due to cleanup errors
                System.Diagnostics.Debug.WriteLine($"Cleanup error: {ex.Message}");
            }
        }
    }
}