using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;

        // A single reusable auto-close timer. Previously each UpdateProgress/ShowSuccess/ShowError
        // created a fresh Timer that was never disposed (a per-export handle leak), and the timer
        // could fire after StatusService disposed the form (Close on a disposed form). One shared,
        // disposed, guarded timer fixes both.
        private Timer _closeTimer;

        public ProgressForm()
        {
            InitializeComponent();
        }

        private void ScheduleClose(int delayMs)
        {
            if (IsDisposed || Disposing)
            {
                return;
            }

            if (_closeTimer == null)
            {
                _closeTimer = new Timer();
                _closeTimer.Tick += CloseTimer_Tick;
            }

            _closeTimer.Stop();
            _closeTimer.Interval = delayMs;
            _closeTimer.Start();
        }

        private void CloseTimer_Tick(object sender, EventArgs e)
        {
            _closeTimer.Stop();
            if (!IsDisposed && !Disposing)
            {
                this.Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _closeTimer != null)
            {
                _closeTimer.Stop();
                _closeTimer.Dispose();
                _closeTimer = null;
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Create progress bar
            this.progressBar = new ProgressBar();
            this.progressBar.Minimum = 0;
            this.progressBar.Maximum = 100;
            this.progressBar.Step = 1;
            this.progressBar.Location = new Point(12, 50);
            this.progressBar.Size = new Size(350, 30);

            // Create status label
            this.lblStatus = new Label();
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new Point(12, 20);
            this.lblStatus.Size = new Size(350, 20);
            this.lblStatus.Text = "Processing...";

            // Configure form
            this.ClientSize = new Size(374, 100);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "SlingMD";
            this.TopMost = true;

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        public void UpdateProgress(string message, int percentage)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string, int>(UpdateProgress), new object[] { message, percentage });
                return;
            }

            this.lblStatus.Text = message;
            this.progressBar.Value = Math.Max(0, Math.Min(100, percentage));

            // Auto-close if we reach 100%
            if (percentage >= 100)
            {
                ScheduleClose(1000);
            }

            this.Refresh();
        }

        public void ShowSuccess(string message, bool autoClose = true)
        {
            UpdateProgress(message, 100);
            this.BackColor = Color.FromArgb(220, 255, 220);

            if (autoClose)
            {
                // Overrides the 1s timer that UpdateProgress(…,100) scheduled (single shared timer).
                ScheduleClose(2000);
            }
        }

        public void ShowError(string message, bool autoClose = false)
        {
            UpdateProgress(message, 100);
            this.BackColor = Color.FromArgb(255, 220, 220);

            if (autoClose)
            {
                ScheduleClose(3000);
            }
            else
            {
                // Keep the error visible: cancel the 1s auto-close that UpdateProgress scheduled.
                _closeTimer?.Stop();
            }
        }
    }
} 