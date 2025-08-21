using System;
using System.Drawing;
using System.Windows.Forms;

namespace BuildingBlocksManager
{
    public partial class ImportAnalysisDialog : Form
    {
        public enum UserChoice
        {
            None,
            ImportOnlyChanged,
            ImportAllAsRequested,
            Cancel
        }

        public UserChoice Choice { get; private set; } = UserChoice.None;
        public BuildingBlockLedger.ChangeAnalysis Analysis { get; private set; }

        private Button btnImportChanged;
        private Button btnImportAll;
        private Button btnCancel;
        private Label lblTitle;
        private Label lblSummary;
        private TextBox txtDetails;

        public ImportAnalysisDialog(BuildingBlockLedger.ChangeAnalysis analysis, string originalRequestType)
        {
            Analysis = analysis;
            InitializeComponent();
            PopulateContent(analysis, originalRequestType);
        }

        private void InitializeComponent()
        {
            this.Text = "Import Analysis Results";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title
            lblTitle = new Label
            {
                Text = "Ledger Analysis Complete",
                Location = new Point(20, 20),
                Size = new Size(450, 25),
                Font = new Font(Font.FontFamily, 12, FontStyle.Bold),
                ForeColor = Color.DarkBlue
            };

            // Summary
            lblSummary = new Label
            {
                Location = new Point(20, 55),
                Size = new Size(450, 60),
                Font = new Font(Font.FontFamily, 9),
                ForeColor = Color.Black
            };

            // Details text box
            txtDetails = new TextBox
            {
                Location = new Point(20, 125),
                Size = new Size(450, 180),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Consolas", 9),
                BackColor = Color.WhiteSmoke
            };

            // Buttons
            btnImportChanged = new Button
            {
                Text = "Import Only Changed",
                Location = new Point(20, 320),
                Size = new Size(160, 30),
                BackColor = Color.LightGreen,
                UseVisualStyleBackColor = false,
                Font = new Font(Font.FontFamily, 8, FontStyle.Bold)
            };
            btnImportChanged.Click += BtnImportChanged_Click;

            btnImportAll = new Button
            {
                Text = "Import All As Requested",
                Location = new Point(190, 320),
                Size = new Size(160, 30),
                BackColor = Color.LightBlue,
                UseVisualStyleBackColor = false,
                Font = new Font(Font.FontFamily, 8)
            };
            btnImportAll.Click += BtnImportAll_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(360, 320),
                Size = new Size(100, 30),
                BackColor = Color.LightCoral,
                UseVisualStyleBackColor = false
            };
            btnCancel.Click += BtnCancel_Click;

            // Add controls to form
            this.Controls.AddRange(new Control[]
            {
                lblTitle, lblSummary, txtDetails,
                btnImportChanged, btnImportAll, btnCancel
            });

            // Set default button and cancel button
            this.AcceptButton = btnImportChanged; // Default to recommended option
            this.CancelButton = btnCancel;
        }

        private void PopulateContent(BuildingBlockLedger.ChangeAnalysis analysis, string originalRequestType)
        {
            // Set summary text with proper line breaks
            lblSummary.Text = analysis.GetSummary().Replace("|", "\n");

            // Build concise details
            var details = "";

            if (analysis.TotalChangedFiles == 0)
            {
                details = "All files are up-to-date.\n\n";
                details += "• No changes detected since last import\n";
                details += "• Consider canceling unless you need to reimport everything\n";
                
                btnImportChanged.Text = "No Import Needed";
                btnImportChanged.BackColor = Color.LightGray;
            }
            else
            {
                details = "CHANGES DETECTED:\n\n";
                
                if (analysis.NewFiles.Count > 0)
                {
                    details += $"• {analysis.NewFiles.Count} new files ready to import\n";
                }
                
                if (analysis.ModifiedFiles.Count > 0)
                {
                    details += $"• {analysis.ModifiedFiles.Count} files updated since last import\n";
                }
                
                if (analysis.UnchangedFiles.Count > 0)
                {
                    details += $"• {analysis.UnchangedFiles.Count} files unchanged\n";
                }
                
                details += "\nRECOMMENDATION:\n";
                details += $"Import only the {analysis.TotalChangedFiles} changed files to save time.";
            }
            
            if (analysis.RemovedEntries.Count > 0)
            {
                details += $"\n\nNOTE: {analysis.RemovedEntries.Count} Building Blocks in template are no longer found in source files:\n";
                foreach (var entry in analysis.RemovedEntries)
                {
                    details += $"• {entry.Name} ({entry.Category})\n";
                }
            }

            txtDetails.Text = details;

            // Update button text to show counts
            if (analysis.TotalChangedFiles > 0)
            {
                btnImportChanged.Text = $"Import Only Changed ({analysis.TotalChangedFiles} files)";
            }
            
            // Make the "Import All" button text clearer for selected file operations
            if (originalRequestType.Contains("Selected"))
            {
                btnImportAll.Text = $"Import All Selected ({analysis.TotalFiles} files)";
            }
            else
            {
                btnImportAll.Text = $"Import All ({analysis.TotalFiles} files)";
            }
        }

        private void BtnImportChanged_Click(object sender, EventArgs e)
        {
            Choice = UserChoice.ImportOnlyChanged;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnImportAll_Click(object sender, EventArgs e)
        {
            Choice = UserChoice.ImportAllAsRequested;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            Choice = UserChoice.Cancel;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}