using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using SlingMD.Outlook.Forms;
using Xunit;

namespace SlingMD.Tests.Forms
{
    /// <summary>
    /// Covers <see cref="BatchFolderPickerForm"/>, which now backs both the batch sling and the
    /// opt-in single-email folder prompt. The single-email path made an unconfigured vault far
    /// easier to hit, so the guards around a blank or relative Inbox path are exercised here.
    /// </summary>
    // WinForms test: VSTest on Windows only. Mono runs skip it with -notrait "Category=WindowsOnly".
    [Trait("Category", "WindowsOnly")]
    public class BatchFolderPickerFormTests
    {
        /// <summary>
        /// Runs the given action on a dedicated STA thread (WinForms controls require STA).
        /// Re-throws any exception raised on that thread so the test fails normally.
        /// </summary>
        private static void RunSta(Action action)
        {
            System.Exception thrown = null;
            Thread staThread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (System.Exception ex)
                {
                    thrown = ex;
                }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            if (thrown != null)
            {
                throw new System.Exception("STA thread failed: " + thrown, thrown);
            }
        }

        private static T GetControl<T>(BatchFolderPickerForm form, string fieldName) where T : Control
        {
            FieldInfo field = typeof(BatchFolderPickerForm).GetField(
                fieldName,
                BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(field);
            return (T)field.GetValue(form);
        }

        private static void InvokeClick(BatchFolderPickerForm form, string handlerName)
        {
            MethodInfo handler = typeof(BatchFolderPickerForm).GetMethod(
                handlerName,
                BindingFlags.Instance | BindingFlags.NonPublic);
            Assert.NotNull(handler);
            handler.Invoke(form, new object[] { form, EventArgs.Empty });
        }

        private static string NewTempInbox()
        {
            string dir = Path.Combine(
                Path.GetTempPath(),
                "SlingMDTests",
                "Picker_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            return dir;
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void UnconfiguredInboxPath_DisablesFolderCreation(string inboxPath)
        {
            RunSta(() =>
            {
                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inboxPath))
                {
                    Assert.False(GetControl<Button>(form, "_btnOk").Enabled);
                    Assert.False(GetControl<ListBox>(form, "_lstFolders").Enabled);
                    Assert.False(GetControl<TextBox>(form, "_txtNewFolder").Enabled);
                }
            });
        }

        [Fact]
        public void RelativeInboxPath_DisablesFolderCreation()
        {
            // A blank VaultBasePath makes GetInboxPath() return a relative path, which would
            // otherwise resolve against Outlook's working directory and create vault folders
            // next to OUTLOOK.EXE.
            RunSta(() =>
            {
                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, Path.Combine("Logic", "Inbox")))
                {
                    Assert.False(GetControl<Button>(form, "_btnOk").Enabled);
                }
            });
        }

        [Fact]
        public void UnconfiguredInboxPath_OkHandlerDoesNotThrow()
        {
            // The AcceptButton can fire the handler even while the button is disabled; the guard
            // must hold, and Result must stay Cancel so the caller aborts the sling.
            RunSta(() =>
            {
                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, string.Empty))
                {
                    InvokeClick(form, "BtnOk_Click");
                    Assert.Equal(BatchFolderPickerForm.PickerResult.Cancel, form.Result);
                    Assert.Equal(string.Empty, form.SelectedFolderPath);
                }
            });
        }

        [Fact]
        public void ValidInboxPath_ListsExistingSubfoldersAndEnablesOk()
        {
            RunSta(() =>
            {
                string inbox = NewTempInbox();
                Directory.CreateDirectory(Path.Combine(inbox, "Projects"));
                Directory.CreateDirectory(Path.Combine(inbox, "Clients"));

                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inbox))
                {
                    Assert.True(GetControl<Button>(form, "_btnOk").Enabled);

                    ListBox list = GetControl<ListBox>(form, "_lstFolders");
                    Assert.Equal(2, list.Items.Count);
                    Assert.Equal("Clients", list.Items[0]);   // sorted, case-insensitive
                    Assert.Equal("Projects", list.Items[1]);
                }

                Directory.Delete(inbox, true);
            });
        }

        [Fact]
        public void TypedNewFolderName_IsCreatedUnderInbox()
        {
            RunSta(() =>
            {
                string inbox = NewTempInbox();

                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inbox))
                {
                    GetControl<TextBox>(form, "_txtNewFolder").Text = "Q3 Renewals";
                    InvokeClick(form, "BtnOk_Click");

                    Assert.Equal(BatchFolderPickerForm.PickerResult.UseSubfolder, form.Result);
                    Assert.Equal(Path.Combine(inbox, "Q3 Renewals"), form.SelectedFolderPath);
                    Assert.True(Directory.Exists(form.SelectedFolderPath));
                    Assert.True(form.CreatedNewFolder);
                }

                Directory.Delete(inbox, true);
            });
        }

        [Fact]
        public void TypedNameMatchingExistingFolder_IsNotReportedAsNewlyCreated()
        {
            // The caller deletes a folder it created when the sling writes nothing. Reporting a
            // pre-existing folder as "created" would make that cleanup delete the user's own folder.
            RunSta(() =>
            {
                string inbox = NewTempInbox();
                Directory.CreateDirectory(Path.Combine(inbox, "Projects"));

                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inbox))
                {
                    GetControl<TextBox>(form, "_txtNewFolder").Text = "Projects";
                    InvokeClick(form, "BtnOk_Click");

                    Assert.Equal(BatchFolderPickerForm.PickerResult.UseSubfolder, form.Result);
                    Assert.False(form.CreatedNewFolder);
                    Assert.True(Directory.Exists(form.SelectedFolderPath));
                }

                Directory.Delete(inbox, true);
            });
        }

        [Fact]
        public void SelectingExistingFolderFromList_IsNotReportedAsNewlyCreated()
        {
            RunSta(() =>
            {
                string inbox = NewTempInbox();
                Directory.CreateDirectory(Path.Combine(inbox, "Clients"));

                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inbox))
                {
                    GetControl<ListBox>(form, "_lstFolders").SelectedIndex = 0;
                    InvokeClick(form, "BtnOk_Click");

                    Assert.Equal(Path.Combine(inbox, "Clients"), form.SelectedFolderPath);
                    Assert.False(form.CreatedNewFolder);
                }

                Directory.Delete(inbox, true);
            });
        }

        [Fact]
        public void Skip_ReportsNoNewlyCreatedFolder()
        {
            RunSta(() =>
            {
                string inbox = NewTempInbox();

                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inbox))
                {
                    InvokeClick(form, "BtnSkip_Click");
                    Assert.False(form.CreatedNewFolder);
                }

                Directory.Delete(inbox, true);
            });
        }

        /// <summary>
        /// Exercises the name validator directly. Driving BtnOk_Click with a bad name is not an
        /// option here: the rejection path raises a modal MessageBox, which would hang the run.
        /// </summary>
        private static bool InvokeIsValidFolderName(string name)
        {
            MethodInfo method = typeof(BatchFolderPickerForm).GetMethod(
                "IsValidFolderName",
                BindingFlags.Static | BindingFlags.NonPublic);
            Assert.NotNull(method);
            return (bool)method.Invoke(null, new object[] { name });
        }

        [Theory]
        [InlineData("..")]
        [InlineData(".")]
        [InlineData("...")]
        [InlineData("CON")]
        [InlineData("lpt9")]
        [InlineData("NUL.txt")]
        [InlineData("sub\\nested")]
        [InlineData("sub/nested")]
        [InlineData("trailing.")]
        [InlineData("trailing ")]
        [InlineData(" leading")]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void UnsafeFolderName_IsRejected(string folderName)
        {
            Assert.False(InvokeIsValidFolderName(folderName));
        }

        [Theory]
        [InlineData("Projects")]
        [InlineData("Q3 Renewals")]
        [InlineData("2026-08 Archive")]
        [InlineData("notes.and.dots")]
        [InlineData("CONTRACTS")]
        public void SafeFolderName_IsAccepted(string folderName)
        {
            Assert.True(InvokeIsValidFolderName(folderName));
        }

        [Fact]
        public void Skip_ReportsUseDefaultWithNoFolder()
        {
            RunSta(() =>
            {
                string inbox = NewTempInbox();

                using (BatchFolderPickerForm form = new BatchFolderPickerForm(1, inbox))
                {
                    InvokeClick(form, "BtnSkip_Click");

                    Assert.Equal(BatchFolderPickerForm.PickerResult.UseDefault, form.Result);
                    Assert.Equal(string.Empty, form.SelectedFolderPath);
                }

                Directory.Delete(inbox, true);
            });
        }

        [Fact]
        public void SingleEmail_UsesSingularTitle()
        {
            RunSta(() =>
            {
                string inbox = NewTempInbox();

                using (BatchFolderPickerForm one = new BatchFolderPickerForm(1, inbox))
                using (BatchFolderPickerForm many = new BatchFolderPickerForm(4, inbox))
                {
                    Assert.Equal("Sling Email", one.Text);
                    Assert.Equal("Sling Multiple Emails", many.Text);
                }

                Directory.Delete(inbox, true);
            });
        }
    }
}
