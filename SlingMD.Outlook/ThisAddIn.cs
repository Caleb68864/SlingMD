using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Ribbon;

namespace SlingMD.Outlook
{
    public partial class ThisAddIn
    {
        private ObsidianSettings _settings;
        private EmailProcessor _emailProcessor;
        private SlingRibbon _ribbon;

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new SlingRibbon(this);
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            bool isFirstLaunchAfterInstall;

            _settings = LoadSettings(out isFirstLaunchAfterInstall);
            _emailProcessor = new EmailProcessor(_settings);

            if (isFirstLaunchAfterInstall && !_settings.HasShownSupportPrompt)
            {
                ShowFirstRunSupportPrompt();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Save settings when Outlook is closing
            if (_settings != null)
            {
                _settings.Save();
            }
        }

        private ObsidianSettings LoadSettings(out bool isFirstLaunchAfterInstall)
        {
            ObsidianSettings settings = new ObsidianSettings();
            isFirstLaunchAfterInstall = !settings.HasSavedSettings();
            settings.Load();
            return settings;
        }

        private void ShowFirstRunSupportPrompt()
        {
            SupportService.ShowBuyMeACoffeePrompt();
            _settings.HasShownSupportPrompt = true;

            try
            {
                _settings.Save();
            }
            catch (ArgumentException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
            catch (IOException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
        }

        private void ShowFirstRunStateSaveWarning(string errorMessage)
        {
            MessageBox.Show(
                "SlingMD showed the support prompt but could not save the first-run state."
                    + Environment.NewLine
                    + Environment.NewLine
                    + errorMessage,
                "SlingMD",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public async void ProcessSelectedEmail()
        {
            try
            {
                // Get selected email
                Explorer explorer = Application.ActiveExplorer();
                if (explorer.Selection.Count == 0)
                {
                    MessageBox.Show("Please select an email first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                MailItem mail = explorer.Selection[1] as MailItem;
                if (mail == null)
                {
                    MessageBox.Show("Selected item is not an email.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Process the email
                await _emailProcessor.ProcessEmail(mail);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving email: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ShowSettings()
        {
            try
            {
                using (SettingsForm form = new SettingsForm(_settings))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Settings are automatically saved by the form
                        // Recreate email processor with new settings
                        _emailProcessor = new EmailProcessor(_settings);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
