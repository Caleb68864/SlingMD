using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace SlingMD.Outlook.Ribbon
{
    [ComVisible(true)]
    public class SlingRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;
        private readonly ThisAddIn _addIn;
        private Bitmap _slingLogo;

        public SlingRibbon(ThisAddIn addIn)
        {
            _addIn = addIn;
            LoadSlingLogo();
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
            try
            {
                return _addIn.GetSelectedItemLabel();
            }
            catch
            {
                return "Sling";
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
            }
        }

        #endregion
    }
} 