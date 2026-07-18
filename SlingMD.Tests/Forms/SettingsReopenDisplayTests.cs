using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Models;
using SlingMD.Tests.Models;
using Xunit;

namespace SlingMD.Tests.Forms
{
    /// <summary>
    /// Repro attempt for issue #13 remaining report: Contacts, Tasks, Threading, Developer tabs
    /// revert. Unlike the existing test, this reopens a NEW form on the reloaded settings and
    /// asserts the CONTROLS display the saved values (the path the user actually observes).
    /// </summary>
    public class SettingsReopenDisplayTests
    {
        private static void RunSta(Action action)
        {
            System.Exception thrown = null;
            Thread t = new Thread(() =>
            {
                try { action(); }
                catch (System.Exception ex) { thrown = ex; }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            if (thrown != null) throw new System.Exception("STA thread failed: " + thrown, thrown);
        }

        private static void InvokeSave(SettingsForm form)
        {
            MethodInfo handler = typeof(SettingsForm).GetMethod("btnSave_Click", BindingFlags.Instance | BindingFlags.NonPublic);
            handler.Invoke(form, new object[] { form, EventArgs.Empty });
        }

        private static T GetControl<T>(SettingsForm form, string field) where T : Control
        {
            FieldInfo f = typeof(SettingsForm).GetField(field, BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(f);
            return (T)f.GetValue(form);
        }

        [Fact]
        public void ChangedValues_DisplayAfterReopen()
        {
            RunSta(() =>
            {
                string dir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "Reopen_" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "ObsidianSettings.json");

                // First run: default settings written to disk.
                ObsidianSettingsTestable first = new ObsidianSettingsTestable { TestSettingsPath = path };
                first.Save();

                // Open form, flip Contacts/Tasks/Threading/Developer via CONTROLS, save.
                using (SettingsForm form = new SettingsForm(first))
                {
                    GetControl<CheckBox>(form, "chkEnableContactSaving").Checked = false;   // default true
                    GetControl<CheckBox>(form, "chkSearchEntireVaultForContacts").Checked = true; // default false
                    GetControl<CheckBox>(form, "chkCreateObsidianTask").Checked = false;     // default true
                    GetControl<CheckBox>(form, "chkCreateOutlookTask").Checked = true;       // default false
                    GetControl<CheckBox>(form, "chkGroupEmailThreads").Checked = false;      // default true
                    GetControl<CheckBox>(form, "chkShowDevelopmentSettings").Checked = true; // default false
                    GetControl<CheckBox>(form, "chkShowThreadDebug").Checked = true;         // default false
                    InvokeSave(form);
                    Assert.Equal(DialogResult.OK, form.DialogResult);
                }

                // Restart: reload from disk.
                ObsidianSettingsTestable reloaded = new ObsidianSettingsTestable { TestSettingsPath = path };
                reloaded.Load();

                // Reopen the form on reloaded settings and check the CONTROLS.
                using (SettingsForm form2 = new SettingsForm(reloaded))
                {
                    Assert.False(GetControl<CheckBox>(form2, "chkEnableContactSaving").Checked);
                    Assert.True(GetControl<CheckBox>(form2, "chkSearchEntireVaultForContacts").Checked);
                    Assert.False(GetControl<CheckBox>(form2, "chkCreateObsidianTask").Checked);
                    Assert.True(GetControl<CheckBox>(form2, "chkCreateOutlookTask").Checked);
                    Assert.False(GetControl<CheckBox>(form2, "chkGroupEmailThreads").Checked);
                    Assert.True(GetControl<CheckBox>(form2, "chkShowDevelopmentSettings").Checked);
                    Assert.True(GetControl<CheckBox>(form2, "chkShowThreadDebug").Checked);
                }
            });
        }
    }
}
