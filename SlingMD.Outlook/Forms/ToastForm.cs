using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// Small notification shown in the top-right corner. It never takes keyboard focus, stacks below
    /// any toasts already showing, grows to fit its text, pauses while the mouse is over it and closes
    /// when clicked.
    /// </summary>
    public class ToastForm : Form
    {
        private const int ToastWidth = 360;
        private const int MinHeight = 56;
        private const int MaxHeight = 220;
        private const int ScreenMargin = 12;
        private const int Gap = 8;

        // Win32 extended styles: no activation, no taskbar/Alt+Tab entry, always on top. TopMost is set
        // through the style rather than the property, because setting the property activates the window.
        private const int WsExNoActivate = 0x08000000;
        private const int WsExToolWindow = 0x00000080;
        private const int WsExTopMost = 0x00000008;

        // Toasts currently on screen, top to bottom; UI thread only.
        private static readonly System.Collections.Generic.List<ToastForm> OpenToasts = new System.Collections.Generic.List<ToastForm>();

        private readonly int _durationMs;
        private Timer _timer;
        private Label _lblMessage;

        public ToastForm(string message, bool isError = false, int durationMs = 4000)
        {
            // Longer messages stay up longer: about 60 ms per character past the first 80.
            int length = message?.Length ?? 0;
            _durationMs = Math.Max(durationMs, durationMs + (length - 80) * 60);
            InitializeComponent(message, isError);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WsExNoActivate | WsExToolWindow | WsExTopMost;
                return cp;
            }
        }

        private void InitializeComponent(string message, bool isError)
        {
            this.SuspendLayout();

            _lblMessage = new Label();
            _lblMessage.AutoSize = false;
            _lblMessage.Dock = DockStyle.Fill;
            _lblMessage.TextAlign = ContentAlignment.MiddleLeft;
            _lblMessage.Font = SystemFonts.MessageBoxFont;
            _lblMessage.ForeColor = Color.White;
            _lblMessage.Text = message;
            _lblMessage.Padding = new Padding(10, 8, 10, 8);
            _lblMessage.AutoEllipsis = true;

            this.Controls.Add(_lblMessage);

            Size textSize = TextRenderer.MeasureText(
                message ?? string.Empty,
                _lblMessage.Font,
                new Size(ToastWidth - _lblMessage.Padding.Horizontal, 0),
                TextFormatFlags.WordBreak);
            int height = Math.Max(MinHeight, Math.Min(MaxHeight, textSize.Height + _lblMessage.Padding.Vertical + 4));

            this.ClientSize = new Size(ToastWidth, height);
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = isError ? Color.FromArgb(180, 60, 60) : Color.FromArgb(50, 50, 50);
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.Cursor = Cursors.Hand;

            _timer = new Timer();
            _timer.Interval = Math.Max(1, _durationMs);
            _timer.Tick += Timer_Tick;

            // Hovering pauses the countdown so a long message can be read; clicking dismisses.
            _lblMessage.MouseEnter += (sender, e) => _timer.Stop();
            _lblMessage.MouseLeave += (sender, e) => _timer.Start();
            _lblMessage.Click += (sender, e) => this.Close();

            this.ResumeLayout(false);
            this.Load += ToastForm_Load;
        }

        private void ToastForm_Load(object sender, EventArgs e)
        {
            OpenToasts.Add(this);
            LayoutOpenToasts();
            _timer.Start();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            OpenToasts.Remove(this);
            LayoutOpenToasts();
            base.OnFormClosed(e);
        }

        /// <summary>Stacks the open toasts downward from the top-right corner of the primary screen.</summary>
        private static void LayoutOpenToasts()
        {
            Rectangle workingArea = Screen.PrimaryScreen.WorkingArea;
            int top = workingArea.Top + ScreenMargin;
            foreach (ToastForm toast in OpenToasts)
            {
                if (toast.IsDisposed)
                {
                    continue;
                }

                toast.Location = new Point(workingArea.Right - toast.Width - ScreenMargin, top);
                top += toast.Height + Gap;
            }
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

        // A marshaling control created on the add-in's UI thread at startup. Toasts are shown from
        // background monitor continuations that may resume off the UI thread; this gives a reliable
        // way to marshal onto the UI thread instead of guessing via Application.OpenForms[0] (which
        // may be empty or a form that is mid-close).
        private static Control _uiMarshaler;

        /// <summary>Registers the UI-thread marshaling control (called once at add-in startup).</summary>
        public static void SetUiMarshaler(Control control)
        {
            _uiMarshaler = control;
        }

        /// <summary>
        /// Shows a non-blocking toast notification. Safe to call from any thread.
        /// </summary>
        public static void ShowToast(string message, bool isError = false, int durationMs = 4000)
        {
            try
            {
                Control marshaler = _uiMarshaler;
                if (marshaler != null && !marshaler.IsDisposed && marshaler.IsHandleCreated)
                {
                    if (marshaler.InvokeRequired)
                    {
                        marshaler.BeginInvoke(new Action(() => ShowToastOnUIThread(message, isError, durationMs)));
                    }
                    else
                    {
                        // !InvokeRequired means we are already on the marshaler's UI thread.
                        ShowToastOnUIThread(message, isError, durationMs);
                    }
                    return;
                }

                // No usable marshaler: try an open form on its UI thread.
                if (Application.OpenForms.Count > 0)
                {
                    Form ownerForm = Application.OpenForms[0];
                    if (ownerForm != null && !ownerForm.IsDisposed && ownerForm.IsHandleCreated)
                    {
                        if (ownerForm.InvokeRequired)
                        {
                            ownerForm.BeginInvoke(new Action(() => ShowToastOnUIThread(message, isError, durationMs)));
                        }
                        else
                        {
                            ShowToastOnUIThread(message, isError, durationMs);
                        }
                        return;
                    }
                }

                // Last resort: only construct the Form directly when we can prove we are on a WinForms
                // UI thread. Otherwise (e.g. a background continuation with no marshaler) constructing
                // a Form off-thread is a crash risk, so drop the toast to a log entry instead.
                if (System.Threading.SynchronizationContext.Current is WindowsFormsSynchronizationContext)
                {
                    ShowToastOnUIThread(message, isError, durationMs);
                }
                else
                {
                    Helpers.Logger.Instance.Warning($"ToastForm: no UI-thread available, dropping toast: {message}");
                }
            }
            catch
            {
                // Silently degrade if toast display fails
            }
        }

        private static void ShowToastOnUIThread(string message, bool isError, int durationMs)
        {
            // Reached directly (synchronous path) and via BeginInvoke (deferred). The outer try in
            // ShowToast does not cover the deferred callback, and the constructor touches
            // Screen.PrimaryScreen (null in some headless/RDP sessions), so guard here too.
            try
            {
                ToastForm toast = new ToastForm(message, isError, durationMs);
                toast.Show();
            }
            catch (System.Exception ex)
            {
                Helpers.Logger.Instance.Warning($"ToastForm: could not display toast: {ex.Message}");
            }
        }
    }
}
