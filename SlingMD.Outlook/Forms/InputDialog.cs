using System;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class InputDialog : BaseForm, IDisposable
    {
        private bool _disposed = false;

        private TextBox txtInput;
        private Button btnOK;
        private Button btnCancel;

        public string InputText => txtInput.Text;

        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            InitializeComponent();
            Text = title;

            // Create and configure controls
            var lblPrompt = new Label
            {
                Text = prompt,
                AutoSize = true,
                Location = new System.Drawing.Point(12, 12)
            };

            txtInput = new TextBox
            {
                Location = new System.Drawing.Point(12, lblPrompt.Bottom + 6),
                Width = 360,
                Text = defaultValue
            };

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(txtInput.Right - 160, txtInput.Bottom + 12),
                Width = 75
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(btnOK.Right + 10, btnOK.Top),
                Width = 75
            };

            // Add controls to form
            Controls.AddRange(new Control[] { lblPrompt, txtInput, btnOK, btnCancel });

            // Set form properties
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            AcceptButton = btnOK;
            CancelButton = btnCancel;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            Padding = new Padding(12);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.ClientSize = new System.Drawing.Size(384, 141);
            this.Name = "InputDialog";
            this.ResumeLayout(false);
        }

        protected override void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources
                    if (txtInput != null) txtInput.Dispose();
                    if (btnOK != null) btnOK.Dispose();
                    if (btnCancel != null) btnCancel.Dispose();
                }

                _disposed = true;
            }
            base.Dispose(disposing);
        }

        ~InputDialog()
        {
            Dispose(false);
        }
    }
} 