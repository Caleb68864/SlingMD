using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public class ToastForm : Form
    {
        private readonly int _durationMs;
        private Timer _timer;
        private Label _lblMessage;

        public ToastForm(string message, bool isError = false, int durationMs = 4000)
        {
            _durationMs = durationMs;
            InitializeComponent(message, isError);
        }

        private void InitializeComponent(string message, bool isError)
        {
            this.SuspendLayout();

            _lblMessage = new Label();
            _lblMessage.AutoSize = false;
            _lblMessage.Dock = DockStyle.Fill;
            _lblMessage.TextAlign = ContentAlignment.MiddleCenter;
            _lblMessage.Font = SystemFonts.MessageBoxFont;
            _lblMessage.ForeColor = Color.White;
            _lblMessage.Text = message;
            _lblMessage.Padding = new Padding(8);

            this.Controls.Add(_lblMessage);

            this.ClientSize = new Size(350, 80);
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = isError ? Color.FromArgb(180, 60, 60) : Color.FromArgb(50, 50, 50);
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;

            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            this.Location = new Point(
                workingArea.Right - this.Width - 12,
                workingArea.Top + 12);

            _timer = new Timer();
            _timer.Interval = _durationMs;
            _timer.Tick += Timer_Tick;

            this.ResumeLayout(false);
            this.Load += ToastForm_Load;
        }

        private void ToastForm_Load(object sender, EventArgs e)
        {
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _timer.Stop();
            this.Close();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && _timer != null)
            {
                _timer.Stop();
                _timer.Dispose();
            }

            base.Dispose(disposing);
        }

        /// <summary>
        /// Shows a non-blocking toast notification. Safe to call from any thread.
        /// </summary>
        public static void ShowToast(string message, bool isError = false, int durationMs = 4000)
        {
            try
            {
                Form ownerForm = null;
                if (Application.OpenForms.Count > 0)
                {
                    ownerForm = Application.OpenForms[0];
                }

                if (ownerForm != null && ownerForm.InvokeRequired)
                {
                    ownerForm.BeginInvoke(new Action(() => ShowToastOnUIThread(message, isError, durationMs)));
                }
                else
                {
                    ShowToastOnUIThread(message, isError, durationMs);
                }
            }
            catch
            {
                // Silently degrade if toast display fails
            }
        }

        private static void ShowToastOnUIThread(string message, bool isError, int durationMs)
        {
            ToastForm toast = new ToastForm(message, isError, durationMs);
            toast.Show();
        }
    }
}
