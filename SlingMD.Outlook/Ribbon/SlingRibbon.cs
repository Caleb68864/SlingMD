using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using SlingMD.Outlook.Forms;

namespace SlingMD.Outlook.Ribbon
{
    [ComVisible(true)]
    public class SlingRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private readonly ThisAddIn _addIn;
        private Bitmap _slingLogo;
        private Bitmap _completeThreadIcon;
        private string _slingButtonLabel = "Sling";

        public SlingRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
            LoadSlingLogo();
            BuildCompleteThreadIcon();
        }

        private void LoadSlingLogo()
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("SlingMD.Outlook.Resources.SlingMD_pixel.png"))
                {
                    if (stream != null)
                    {
                        _slingLogo = new Bitmap(stream);
                    }
                }
            }
            catch (Exception)
            {
                // If loading fails, we'll fall back to the default Office icon
                _slingLogo = null;
            }
        }

        /// <summary>
        /// Builds the Complete Thread button image: the Sling logo with a green check
        /// badge composited in the bottom-right corner. Cached once at construction.
        /// Falls back to the unmodified Sling logo if compositing fails.
        /// </summary>
        private void BuildCompleteThreadIcon()
        {
            if (_slingLogo == null) return;

            try
            {
                Bitmap composed = new Bitmap(_slingLogo.Width, _slingLogo.Height);
                using (Graphics g = Graphics.FromImage(composed))
                {
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    g.DrawImage(_slingLogo, 0, 0, _slingLogo.Width, _slingLogo.Height);

                    // Green check badge, ~40% of the icon, positioned bottom-right.
                    int badgeSize = (int)(Math.Min(composed.Width, composed.Height) * 0.55);
                    int padding = Math.Max(1, composed.Width / 32);
                    int x = composed.Width - badgeSize - padding;
                    int y = composed.Height - badgeSize - padding;

                    // White halo so the badge stays legible against dark sling shapes.
                    using (var halo = new SolidBrush(Color.White))
                    {
                        g.FillEllipse(halo, x - 1, y - 1, badgeSize + 2, badgeSize + 2);
                    }

                    // Filled green disc.
                    using (var fill = new SolidBrush(Color.FromArgb(46, 160, 67)))
                    {
                        g.FillEllipse(fill, x, y, badgeSize, badgeSize);
                    }

                    // Checkmark: two-segment polyline centered in the disc.
                    using (var pen = new Pen(Color.White, Math.Max(1f, badgeSize / 8f))
                    {
                        StartCap = System.Drawing.Drawing2D.LineCap.Round,
                        EndCap = System.Drawing.Drawing2D.LineCap.Round,
                        LineJoin = System.Drawing.Drawing2D.LineJoin.Round
                    })
                    {
                        PointF p1 = new PointF(x + badgeSize * 0.25f, y + badgeSize * 0.53f);
                        PointF p2 = new PointF(x + badgeSize * 0.44f, y + badgeSize * 0.72f);
                        PointF p3 = new PointF(x + badgeSize * 0.76f, y + badgeSize * 0.32f);
                        g.DrawLines(pen, new[] { p1, p2, p3 });
                    }
                }
                _completeThreadIcon = composed;
            }
            catch (Exception)
            {
                _completeThreadIcon = _slingLogo;
            }
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Appointment")
            {
                return @"<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon><tabs>
    <tab idMso=""TabAppointment"">
      <group id=""SlingAppointmentGroup"" label=""Sling"" insertBeforeMso=""GroupActions"">
        <button id=""InspectorSlingButton"" label=""Sling"" size=""large""
                getImage=""GetSlingButtonImage"" onAction=""OnInspectorSlingClick""
                supertip=""Save this appointment to Obsidian as a markdown note""/>
      </group>
    </tab>
  </tabs></ribbon>
</customUI>";
            }

            return GetResourceText("SlingMD.Outlook.Ribbon.SlingRibbon.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void InvalidateSlingButton()
        {
            _ribbon?.InvalidateControl("SlingButton");
        }

        #endregion

        #region Ribbon Callbacks

        public void OnSlingButtonClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.ProcessSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing item: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSaveTodaysClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.SaveTodaysAppointments();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving today's appointments: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSaveDateRangeClick(Office.IRibbonControl control)
        {
            try
            {
                using (CalendarRangeDialog dialog = new CalendarRangeDialog())
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        _addIn.SaveAppointmentRange(dialog.StartDate, dialog.EndDate);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnInspectorSlingClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.ProcessCurrentAppointment();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error processing appointment: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public string GetSlingButtonLabel(Office.IRibbonControl control)
        {
            return _slingButtonLabel;
        }

        public void UpdateSlingButtonLabel(string label)
        {
            _slingButtonLabel = label ?? "Sling";
            _ribbon?.InvalidateControl("SlingButton");
        }

        public void OnCompleteThreadClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.CompleteThread();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error completing thread: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSlingAllContactsClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.SlingAllContacts();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting contacts: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnSettingsButtonClick(Office.IRibbonControl control)
        {
            try
            {
                _addIn.ShowSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public Bitmap GetSlingButtonImage(Office.IRibbonControl control)
        {
            return _slingLogo;
        }

        public Bitmap GetCompleteThreadImage(Office.IRibbonControl control)
        {
            return _completeThreadIcon ?? _slingLogo;
        }

        #endregion

        #region Helpers

        private string GetResourceText(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _slingLogo?.Dispose();
                _completeThreadIcon?.Dispose();
            }
        }

        #endregion
    }
} 