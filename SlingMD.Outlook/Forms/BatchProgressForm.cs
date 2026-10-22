using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// Non-modal progress window for batch slings ("Slinging 3 of 25: subject") with a Cancel button.
    /// Cancelling stops the batch after the email in progress; it never interrupts one mid-write.
    /// </summary>
    public class BatchProgressForm : Form
    {
        private readonly int _total;
        private Label lblStatus;
        private Label lblSubject;
        private ProgressBar progressBar;
        private Button btnCancel;

        public BatchProgressForm(string title, int total)
        {
            _total = Math.Max(total, 1);
            InitializeComponent();
            this.Text = string.IsNullOrEmpty(title) ? "SlingMD" : title;
        }

        /// <summary>Whether the user clicked Cancel (or closed the window).</summary>
        public bool CancelRequested { get; private set; }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.lblStatus = new Label
            {
                AutoSize = false,
                Location = new Point(12, 12),
                Size = new Size(400, 20),
                Text = "Starting..."
            };

            this.lblSubject = new Label
            {
                AutoSize = false,
                AutoEllipsis = true,
                Location = new Point(12, 34),
                Size = new Size(400, 20),
                ForeColor = SystemColors.GrayText
            };

            this.progressBar = new ProgressBar
            {
                Minimum = 0,
                Maximum = _total,
                Location = new Point(12, 60),
                Size = new Size(400, 24)
            };

            this.btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(337, 94),
                Size = new Size(75, 26)
            };
            this.btnCancel.Click += (sender, e) => RequestCancel();

            this.ClientSize = new Size(424, 132);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnCancel);
            this.CancelButton = this.btnCancel;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;

            this.ResumeLayout(false);
        }

        /// <summary>Shows which email is being processed (1-based <paramref name="index"/>).</summary>
        public void ShowItem(int index, string subject)
        {
            if (IsDisposed)
            {
                return;
            }

            lblStatus.Text = CancelRequested
                ? "Cancelling after this email..."
                : $"Slinging {index} of {_total}";
            lblSubject.Text = subject ?? string.Empty;
            progressBar.Value = Math.Max(0, Math.Min(_total, index - 1));
        }

        private void RequestCancel()
        {
            CancelRequested = true;
            btnCancel.Enabled = false;
            lblStatus.Text = "Cancelling after this email...";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            // Closing the window by its X counts as Cancel; the batch closes it when done.
            if (e.CloseReason == CloseReason.UserClosing && !CancelRequested)
            {
                RequestCancel();
            }
            base.OnFormClosing(e);
        }
    }
}
