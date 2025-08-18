using System;
using System.Windows.Forms;

namespace BuildingBlocksManager
{
    public partial class LoggingConfigForm : Form
    {
        private Settings settings;
        private CheckBox chkLogToTemplateDirectory;
        private CheckBox chkEnableDetailedLogging;
        private Button btnOK;
        private Button btnCancel;
        private Label lblLogLocation;

        public LoggingConfigForm(Settings settings)
        {
            this.settings = settings;
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Logging Configuration";
            this.Size = new System.Drawing.Size(450, 220);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Log location label
            lblLogLocation = new Label
            {
                Text = "Log Files Location:",
                Location = new System.Drawing.Point(12, 15),
                Size = new System.Drawing.Size(120, 23),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft
            };
            this.Controls.Add(lblLogLocation);

            // Log to template directory checkbox
            chkLogToTemplateDirectory = new CheckBox
            {
                Text = "Store logs in template directory (BBM_Logs folder)",
                Location = new System.Drawing.Point(12, 45),
                Size = new System.Drawing.Size(400, 23),
                Checked = settings.LogToTemplateDirectory
            };
            chkLogToTemplateDirectory.CheckedChanged += ChkLogToTemplateDirectory_CheckedChanged;
            this.Controls.Add(chkLogToTemplateDirectory);

            var lblNote = new Label
            {
                Text = "Note: If unchecked, logs will be stored in the source directory or user profile.",
                Location = new System.Drawing.Point(30, 68),
                Size = new System.Drawing.Size(380, 23),
                ForeColor = System.Drawing.SystemColors.GrayText,
                Font = new System.Drawing.Font("Segoe UI", 8F)
            };
            this.Controls.Add(lblNote);

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
                Size = new System.Drawing.Size(380, 23),
                ForeColor = System.Drawing.SystemColors.GrayText,
                Font = new System.Drawing.Font("Segoe UI", 8F)
            };
            this.Controls.Add(lblDetailNote);

            // OK Button
            btnOK = new Button
            {
                Text = "OK",
                Location = new System.Drawing.Point(270, 160),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // Cancel Button
            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new System.Drawing.Point(355, 160),
                Size = new System.Drawing.Size(75, 23),
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void ChkLogToTemplateDirectory_CheckedChanged(object sender, EventArgs e)
        {
            // Could add preview of where logs would be stored here
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            settings.LogToTemplateDirectory = chkLogToTemplateDirectory.Checked;
            settings.EnableDetailedLogging = chkEnableDetailedLogging.Checked;
            settings.Save();
        }
    }
}