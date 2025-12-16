using System;
using System.IO;
using System.Windows.Forms;

namespace BuildingBlocksManager
{
    public partial class LoggingConfigForm : Form
    {
        private Settings settings;
        private TextBox txtLogDirectory;
        private Button btnBrowseLogDirectory;
        private CheckBox chkEnableDetailedLogging;
        private Button btnOK;
        private Button btnCancel;
        private Label lblLogLocation;
        private Label lblCurrentPath;

        public LoggingConfigForm(Settings settings)
        {
            this.settings = settings;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Logging Configuration";
            this.Size = new System.Drawing.Size(550, 220);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Log location label
            lblLogLocation = new Label
            {
                Text = "Log Directory:",
                Location = new System.Drawing.Point(12, 15),
                Size = new System.Drawing.Size(100, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblLogLocation);

            // Log directory text box
            txtLogDirectory = new TextBox
            {
                Location = new System.Drawing.Point(12, 40),
                Size = new System.Drawing.Size(440, 23),
                Text = settings.CurrentPaths.LogDirectory
            };
            txtLogDirectory.TextChanged += TxtLogDirectory_TextChanged;
            this.Controls.Add(txtLogDirectory);

            // Browse button
            btnBrowseLogDirectory = new Button
            {
                Text = "Browse...",
                Location = new System.Drawing.Point(458, 39),
                Size = new System.Drawing.Size(75, 25)
            };
            btnBrowseLogDirectory.Click += BtnBrowseLogDirectory_Click;
            this.Controls.Add(btnBrowseLogDirectory);

            // Current path display
            lblCurrentPath = new Label
            {
                Text = GetCurrentPathDisplay(),
                Location = new System.Drawing.Point(12, 68),
                Size = new System.Drawing.Size(520, 23),
                ForeColor = System.Drawing.SystemColors.GrayText,
                Font = new System.Drawing.Font("Segoe UI", 8F)
            };
            this.Controls.Add(lblCurrentPath);

            // Enable detailed logging checkbox
            chkEnableDetailedLogging = new CheckBox
            {
                Text = "Enable detailed logging (includes INFO level messages)",
                Location = new System.Drawing.Point(12, 105),
                Size = new System.Drawing.Size(400, 23),
                Checked = settings.EnableDetailedLogging
            };
            this.Controls.Add(chkEnableDetailedLogging);

            var lblDetailNote = new Label
            {
                Text = "Note: When disabled, only WARNING, ERROR, and SUCCESS messages are logged.",
                Location = new System.Drawing.Point(30, 128),
                Size = new System.Drawing.Size(500, 23),
                ForeColor = System.Drawing.SystemColors.GrayText,
                Font = new System.Drawing.Font("Segoe UI", 8F)
            };
            this.Controls.Add(lblDetailNote);

            // OK Button
            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(370, 160),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // Cancel Button
            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(455, 160),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void BtnBrowseLogDirectory_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Log Directory";
                dialog.ShowNewFolderButton = true;

                if (!string.IsNullOrEmpty(txtLogDirectory.Text) && Directory.Exists(txtLogDirectory.Text))
                {
                    dialog.SelectedPath = txtLogDirectory.Text;
                }

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtLogDirectory.Text = dialog.SelectedPath;
                }
            }
        }

        private void TxtLogDirectory_TextChanged(object sender, EventArgs e)
        {
            lblCurrentPath.Text = GetCurrentPathDisplay();
        }

        private string GetCurrentPathDisplay()
        {
            if (string.IsNullOrWhiteSpace(txtLogDirectory.Text))
            {
                return "Logging disabled - please set a log directory";
            }
            else
            {
                return $"Logs will be stored in: {txtLogDirectory.Text}\\BBM_Logs";
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            settings.CurrentPaths.LogDirectory = txtLogDirectory.Text.Trim();
            settings.EnableDetailedLogging = chkEnableDetailedLogging.Checked;
            settings.Save();
        }
    }
}