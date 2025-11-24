using System;
using System.Windows.Forms;
using System.Drawing;
using SlingMD.Outlook.Models;
using System.Collections.Generic;
using System.Reflection;
using System.Linq;

namespace SlingMD.Outlook.Forms
{
    public partial class SettingsForm : Form
    {
        private readonly ObsidianSettings _settings;
        
        // Designer-generated fields
        private TextBox txtVaultName;
        private TextBox txtVaultPath;
        private TextBox txtInboxFolder;
        private TextBox txtContactsFolder;
        private CheckBox chkLaunchObsidian;
        private CheckBox chkEnableContactSaving;
        private CheckBox chkSearchEntireVaultForContacts;
        private NumericUpDown numDelay;
        private CheckBox chkShowCountdown;
        private CheckBox chkCreateObsidianTask;
        private CheckBox chkCreateOutlookTask;
        private CheckBox chkAskForDates;
        private CheckBox chkGroupEmailThreads;
        private NumericUpDown numDefaultDueDays;
        private NumericUpDown numDefaultReminderDays;
        private NumericUpDown numDefaultReminderHour;
        private Button btnBrowse;
        private Button btnSave;
        private Button btnCancel;
        private ListBox lstPatterns;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnRemove;
        private GroupBox grpDevelopment;
        private CheckBox chkShowDevelopmentSettings;
        private CheckBox chkShowThreadDebug;
        // New controls for note/tag customization
        private TextBox txtDefaultNoteTags;
        private TextBox txtDefaultTaskTags;
        private TextBox txtNoteTitleFormat;
        private NumericUpDown numNoteTitleMaxLength;
        private CheckBox chkNoteTitleIncludeDate;
        private Label lblDefaultNoteTags;
        private Label lblDefaultTaskTags;
        private Label lblNoteTitleFormat;
        private Label lblNoteTitleMaxLength;
        private Label lblNoteTitleIncludeDate;
        private GroupBox grpNoteCustomization;
        private ToolTip toolTip;
        // Refactored layout containers
        private TableLayoutPanel mainLayout;
        private CheckBox chkMoveDateToFrontInThread;
        // Attachment controls
        private GroupBox grpAttachments;
        private TextBox txtAttachmentsFolder;
        private ComboBox cmbAttachmentStorageMode;
        private CheckBox chkSaveInlineImages;
        private CheckBox chkSaveAllAttachments;
        private CheckBox chkUseObsidianWikilinks;
        private Label lblAttachmentsFolder;
        private Label lblAttachmentStorageMode;

        public SettingsForm(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.toolTip = new ToolTip();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));

            // Main layout panel
            this.mainLayout = new TableLayoutPanel();
            this.mainLayout.Dock = DockStyle.Fill;
            this.mainLayout.AutoScroll = true;
            this.mainLayout.ColumnCount = 1;
            this.mainLayout.RowCount = 0;
            this.mainLayout.GrowStyle = TableLayoutPanelGrowStyle.AddRows;

            // Group: Vault/Paths
            var grpVault = new GroupBox();
            grpVault.Text = "Vault Settings";
            grpVault.Dock = DockStyle.Top;
            grpVault.AutoSize = true;
            var vaultLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            vaultLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            vaultLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            vaultLayout.Controls.Add(new Label { Text = "Vault Name:", Anchor = AnchorStyles.Left }, 0, 0);
            this.txtVaultName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
            vaultLayout.Controls.Add(this.txtVaultName, 1, 0);
            vaultLayout.Controls.Add(new Label { Text = "Vault Base Path:", Anchor = AnchorStyles.Left }, 0, 1);
            this.txtVaultPath = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
            vaultLayout.Controls.Add(this.txtVaultPath, 1, 1);
            this.btnBrowse = new Button { Text = "Browse...", Anchor = AnchorStyles.Left };
            vaultLayout.Controls.Add(this.btnBrowse, 1, 2);
            vaultLayout.Controls.Add(new Label { Text = "Inbox Folder:", Anchor = AnchorStyles.Left }, 0, 3);
            this.txtInboxFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
            vaultLayout.Controls.Add(this.txtInboxFolder, 1, 3);
            vaultLayout.Controls.Add(new Label { Text = "Contacts Folder:", Anchor = AnchorStyles.Left }, 0, 4);
            this.txtContactsFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 350 };
            vaultLayout.Controls.Add(this.txtContactsFolder, 1, 4);
            grpVault.Controls.Add(vaultLayout);
            this.mainLayout.Controls.Add(grpVault);

            // Group: General Settings (labels and checkboxes grow with width)
            var grpGeneral = new GroupBox();
            grpGeneral.Text = "General Settings";
            grpGeneral.Dock = DockStyle.Top;
            grpGeneral.AutoSize = true;
            var generalLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            generalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            generalLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
            this.chkEnableContactSaving = new CheckBox { Text = "Enable Contact Saving", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkEnableContactSaving, 0, 0);
            this.chkSearchEntireVaultForContacts = new CheckBox { Text = "Search entire vault for contacts", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkSearchEntireVaultForContacts, 1, 0);
            this.chkLaunchObsidian = new CheckBox { Text = "Launch Obsidian after saving", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkLaunchObsidian, 0, 1);
            this.chkShowCountdown = new CheckBox { Text = "Show countdown", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkShowCountdown, 1, 1);
            this.chkCreateObsidianTask = new CheckBox { Text = "Create task in Obsidian note", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkCreateObsidianTask, 0, 2);
            this.chkCreateOutlookTask = new CheckBox { Text = "Create task in Outlook", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkCreateOutlookTask, 1, 2);
            this.chkAskForDates = new CheckBox { Text = "Ask for dates and times each time", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkAskForDates, 0, 3);
            this.chkGroupEmailThreads = new CheckBox { Text = "Group email threads", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalLayout.Controls.Add(this.chkGroupEmailThreads, 1, 3);
            grpGeneral.Controls.Add(generalLayout);
            this.mainLayout.Controls.Add(grpGeneral);

            // Group: Delay/Task timing
            var grpTiming = new GroupBox();
            grpTiming.Text = "Timing Settings";
            grpTiming.Dock = DockStyle.Top;
            grpTiming.AutoSize = true;
            var timingLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            timingLayout.Controls.Add(new Label { Text = "Delay (seconds):", Anchor = AnchorStyles.Left }, 0, 0);
            this.numDelay = new NumericUpDown { Minimum = 0, Maximum = 10, Anchor = AnchorStyles.Left };
            timingLayout.Controls.Add(this.numDelay, 1, 0);
            timingLayout.Controls.Add(new Label { Text = "Due in Days:", Anchor = AnchorStyles.Left }, 0, 1);
            this.numDefaultDueDays = new NumericUpDown { Minimum = 0, Maximum = 30, Anchor = AnchorStyles.Left };
            timingLayout.Controls.Add(this.numDefaultDueDays, 1, 1);
            timingLayout.Controls.Add(new Label { Text = "Reminder Days:", Anchor = AnchorStyles.Left }, 0, 2);
            this.numDefaultReminderDays = new NumericUpDown { Minimum = 0, Maximum = 30, Anchor = AnchorStyles.Left };
            timingLayout.Controls.Add(this.numDefaultReminderDays, 1, 2);
            timingLayout.Controls.Add(new Label { Text = "Reminder Hour:", Anchor = AnchorStyles.Left }, 0, 3);
            this.numDefaultReminderHour = new NumericUpDown { Minimum = 0, Maximum = 23, Anchor = AnchorStyles.Left };
            timingLayout.Controls.Add(this.numDefaultReminderHour, 1, 3);
            grpTiming.Controls.Add(timingLayout);
            this.mainLayout.Controls.Add(grpTiming);

            // Group: Patterns
            var grpPatterns = new GroupBox();
            grpPatterns.Text = "Subject Cleanup Patterns";
            grpPatterns.Dock = DockStyle.Top;
            grpPatterns.AutoSize = true;
            var patternsLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            this.lstPatterns = new ListBox { Height = 120, Width = 400 };
            patternsLayout.Controls.Add(this.lstPatterns, 0, 0);
            var btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true };
            this.btnAdd = new Button { Text = "Add" };
            this.btnEdit = new Button { Text = "Edit" };
            this.btnRemove = new Button { Text = "Remove" };
            btnPanel.Controls.Add(this.btnAdd);
            btnPanel.Controls.Add(this.btnEdit);
            btnPanel.Controls.Add(this.btnRemove);
            patternsLayout.Controls.Add(btnPanel, 1, 0);
            grpPatterns.Controls.Add(patternsLayout);
            this.mainLayout.Controls.Add(grpPatterns);

            // Group: Note/Tag Customization
            this.grpNoteCustomization = new GroupBox();
            this.grpNoteCustomization.Text = "Note & Tag Customization";
            this.grpNoteCustomization.Dock = DockStyle.Top;
            this.grpNoteCustomization.AutoSize = true;
            var noteTagLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            noteTagLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            noteTagLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));
            this.lblDefaultNoteTags = new Label { Text = "Default Note Tags (comma-separated):", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtDefaultNoteTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320, Dock = DockStyle.Fill };
            noteTagLayout.Controls.Add(this.lblDefaultNoteTags, 0, 0);
            noteTagLayout.Controls.Add(this.txtDefaultNoteTags, 1, 0);
            this.lblDefaultTaskTags = new Label { Text = "Default Task Tags (comma-separated, will be rendered as #tags):", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtDefaultTaskTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320, Dock = DockStyle.Fill };
            noteTagLayout.Controls.Add(this.lblDefaultTaskTags, 0, 1);
            noteTagLayout.Controls.Add(this.txtDefaultTaskTags, 1, 1);
            this.lblNoteTitleFormat = new Label { Text = "Note Title Format (use {Subject}, {Sender}, {Date}):", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtNoteTitleFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320, Dock = DockStyle.Fill };
            noteTagLayout.Controls.Add(this.lblNoteTitleFormat, 0, 2);
            noteTagLayout.Controls.Add(this.txtNoteTitleFormat, 1, 2);
            this.lblNoteTitleMaxLength = new Label { Text = "Max Title Length:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.numNoteTitleMaxLength = new NumericUpDown { Minimum = 10, Maximum = 200, Anchor = AnchorStyles.Left, Dock = DockStyle.Fill };
            noteTagLayout.Controls.Add(this.lblNoteTitleMaxLength, 0, 3);
            noteTagLayout.Controls.Add(this.numNoteTitleMaxLength, 1, 3);
            // Add Note Title Include Date and Move Date checkboxes in a horizontal layout for alignment
            var dateOptionsLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Anchor = AnchorStyles.Left };
            this.chkNoteTitleIncludeDate = new CheckBox { Anchor = AnchorStyles.Left };
            this.lblNoteTitleIncludeDate = new Label { Text = "Include Date in Title:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
            dateOptionsLayout.Controls.Add(this.lblNoteTitleIncludeDate);
            dateOptionsLayout.Controls.Add(this.chkNoteTitleIncludeDate);
            this.chkMoveDateToFrontInThread = new CheckBox { Text = "Move date to front of filename when grouping threads", Anchor = AnchorStyles.Left, AutoSize = true };
            dateOptionsLayout.Controls.Add(this.chkMoveDateToFrontInThread);
            noteTagLayout.Controls.Add(dateOptionsLayout, 1, 4);

            // Add event handler for enabling/disabling move date checkbox
            this.chkNoteTitleIncludeDate.CheckedChanged += (s, e) => {
                if (!chkNoteTitleIncludeDate.Checked) {
                    chkMoveDateToFrontInThread.Checked = false;
                    chkMoveDateToFrontInThread.Enabled = false;
                } else {
                    chkMoveDateToFrontInThread.Enabled = true;
                }
            };

            this.grpNoteCustomization.Controls.Add(noteTagLayout);
            this.mainLayout.Controls.Add(this.grpNoteCustomization);

            // Group: Attachment Settings
            this.grpAttachments = new GroupBox();
            this.grpAttachments.Text = "Attachment Settings";
            this.grpAttachments.Dock = DockStyle.Top;
            this.grpAttachments.AutoSize = true;
            var attachmentLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));

            this.lblAttachmentStorageMode = new Label { Text = "Storage Location:", Anchor = AnchorStyles.Left };
            this.cmbAttachmentStorageMode = new ComboBox {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Width = 320
            };
            this.cmbAttachmentStorageMode.Items.Add("Same folder as note");
            this.cmbAttachmentStorageMode.Items.Add("Subfolder per note");
            this.cmbAttachmentStorageMode.Items.Add("Centralized folder");
            attachmentLayout.Controls.Add(this.lblAttachmentStorageMode, 0, 0);
            attachmentLayout.Controls.Add(this.cmbAttachmentStorageMode, 1, 0);

            this.lblAttachmentsFolder = new Label {
                Text = "Centralized Folder Name:",
                Anchor = AnchorStyles.Left
            };
            this.txtAttachmentsFolder = new TextBox {
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Width = 320
            };
            attachmentLayout.Controls.Add(this.lblAttachmentsFolder, 0, 1);
            attachmentLayout.Controls.Add(this.txtAttachmentsFolder, 1, 1);

            var checkboxLayout = new FlowLayoutPanel {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Left
            };
            this.chkSaveInlineImages = new CheckBox {
                Text = "Save inline images",
                AutoSize = true
            };
            this.chkSaveAllAttachments = new CheckBox {
                Text = "Save all attachments",
                AutoSize = true,
                Margin = new Padding(20, 0, 0, 0)
            };
            this.chkUseObsidianWikilinks = new CheckBox {
                Text = "Use Obsidian wikilinks",
                AutoSize = true,
                Margin = new Padding(20, 0, 0, 0)
            };
            checkboxLayout.Controls.Add(this.chkSaveInlineImages);
            checkboxLayout.Controls.Add(this.chkSaveAllAttachments);
            checkboxLayout.Controls.Add(this.chkUseObsidianWikilinks);
            attachmentLayout.Controls.Add(checkboxLayout, 1, 2);

            // Add event handler to enable/disable centralized folder textbox
            this.cmbAttachmentStorageMode.SelectedIndexChanged += (s, e) => {
                this.txtAttachmentsFolder.Enabled = this.cmbAttachmentStorageMode.SelectedIndex == 2;
            };

            this.grpAttachments.Controls.Add(attachmentLayout);
            this.mainLayout.Controls.Add(this.grpAttachments);

            // Group: Development (always at bottom, always visible)
            this.grpDevelopment = new GroupBox();
            this.grpDevelopment.Text = "Development Settings";
            this.grpDevelopment.Dock = DockStyle.Top;
            this.grpDevelopment.AutoSize = true;
            var devLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            this.chkShowDevelopmentSettings = new CheckBox { Text = "Show development settings", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            this.chkShowThreadDebug = new CheckBox { Text = "Show thread debug", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            devLayout.Controls.Add(this.chkShowDevelopmentSettings);
            devLayout.Controls.Add(this.chkShowThreadDebug);
            this.grpDevelopment.Controls.Add(devLayout);
            this.mainLayout.Controls.Add(this.grpDevelopment);

            // Save/Cancel buttons
            var btnLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Bottom, AutoSize = true };
            this.btnSave = new Button { Text = "Save", DialogResult = DialogResult.OK };
            this.btnSave.Click += btnSave_Click;
            this.btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
            btnLayout.Controls.Add(this.btnSave);
            btnLayout.Controls.Add(this.btnCancel);
            this.mainLayout.Controls.Add(btnLayout);

            // Set up the form
            this.Controls.Add(this.mainLayout);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Obsidian Settings";
            this.ClientSize = new System.Drawing.Size(700, 800);
        }

        private void LoadSettings()
        {
            txtVaultName.Text = _settings.VaultName;
            txtVaultPath.Text = _settings.VaultBasePath;
            txtInboxFolder.Text = _settings.InboxFolder;
            txtContactsFolder.Text = _settings.ContactsFolder;
            chkLaunchObsidian.Checked = _settings.LaunchObsidian;
            chkEnableContactSaving.Checked = _settings.EnableContactSaving;
            chkSearchEntireVaultForContacts.Checked = _settings.SearchEntireVaultForContacts;
            txtContactsFolder.Enabled = _settings.EnableContactSaving;
            numDelay.Value = _settings.ObsidianDelaySeconds;
            chkShowCountdown.Checked = _settings.ShowCountdown;
            chkCreateObsidianTask.Checked = _settings.CreateObsidianTask;
            chkCreateOutlookTask.Checked = _settings.CreateOutlookTask;
            chkAskForDates.Checked = _settings.AskForDates;
            chkGroupEmailThreads.Checked = _settings.GroupEmailThreads;
            numDefaultDueDays.Value = _settings.DefaultDueDays;
            numDefaultReminderDays.Value = _settings.DefaultReminderDays;
            numDefaultReminderHour.Value = _settings.DefaultReminderHour;
            chkShowDevelopmentSettings.Checked = _settings.ShowDevelopmentSettings;
            chkShowThreadDebug.Checked = _settings.ShowThreadDebug;
            chkMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            chkMoveDateToFrontInThread.Enabled = chkNoteTitleIncludeDate.Checked;

            // Initialize development settings visibility
            grpDevelopment.Visible = _settings.ShowDevelopmentSettings;
            chkShowThreadDebug.Visible = _settings.ShowDevelopmentSettings;

            // Load patterns
            lstPatterns.Items.Clear();
            foreach (var pattern in _settings.SubjectCleanupPatterns)
            {
                lstPatterns.Items.Add(pattern);
            }

            txtDefaultNoteTags.Text = string.Join(", ", _settings.DefaultNoteTags ?? new List<string>());
            txtDefaultTaskTags.Text = string.Join(", ", _settings.DefaultTaskTags ?? new List<string>());
            txtNoteTitleFormat.Text = _settings.NoteTitleFormat ?? "{Subject} - {Date}";
            numNoteTitleMaxLength.Value = _settings.NoteTitleMaxLength > 0 ? _settings.NoteTitleMaxLength : 50;
            chkNoteTitleIncludeDate.Checked = _settings.NoteTitleIncludeDate;
            chkMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            chkMoveDateToFrontInThread.Enabled = chkNoteTitleIncludeDate.Checked;

            // Load attachment settings
            txtAttachmentsFolder.Text = _settings.AttachmentsFolder ?? "Attachments";
            cmbAttachmentStorageMode.SelectedIndex = (int)_settings.AttachmentStorageMode;
            txtAttachmentsFolder.Enabled = _settings.AttachmentStorageMode == AttachmentStorageMode.Centralized;
            chkSaveInlineImages.Checked = _settings.SaveInlineImages;
            chkSaveAllAttachments.Checked = _settings.SaveAllAttachments;
            chkUseObsidianWikilinks.Checked = _settings.UseObsidianWikilinks;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select Obsidian Vault Base Directory";
                dialog.SelectedPath = txtVaultPath.Text;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    txtVaultPath.Text = dialog.SelectedPath;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            _settings.VaultName = txtVaultName.Text;
            _settings.VaultBasePath = txtVaultPath.Text;
            _settings.InboxFolder = txtInboxFolder.Text;
            _settings.ContactsFolder = txtContactsFolder.Text;
            _settings.LaunchObsidian = chkLaunchObsidian.Checked;
            _settings.EnableContactSaving = chkEnableContactSaving.Checked;
            _settings.SearchEntireVaultForContacts = chkSearchEntireVaultForContacts.Checked;
            _settings.ObsidianDelaySeconds = (int)numDelay.Value;
            _settings.ShowCountdown = chkShowCountdown.Checked;
            _settings.CreateObsidianTask = chkCreateObsidianTask.Checked;
            _settings.CreateOutlookTask = chkCreateOutlookTask.Checked;
            _settings.AskForDates = chkAskForDates.Checked;
            _settings.GroupEmailThreads = chkGroupEmailThreads.Checked;
            _settings.DefaultDueDays = (int)numDefaultDueDays.Value;
            _settings.DefaultReminderDays = (int)numDefaultReminderDays.Value;
            _settings.DefaultReminderHour = (int)numDefaultReminderHour.Value;
            _settings.ShowDevelopmentSettings = chkShowDevelopmentSettings.Checked;
            _settings.ShowThreadDebug = chkShowThreadDebug.Checked;
            _settings.MoveDateToFrontInThread = chkMoveDateToFrontInThread.Checked;

            // Save patterns
            _settings.SubjectCleanupPatterns.Clear();
            foreach (string pattern in lstPatterns.Items)
            {
                _settings.SubjectCleanupPatterns.Add(pattern);
            }

            _settings.DefaultNoteTags = txtDefaultNoteTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)).ToList();
            _settings.DefaultTaskTags = txtDefaultTaskTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)).ToList();
            _settings.NoteTitleFormat = txtNoteTitleFormat.Text.Trim();
            _settings.NoteTitleMaxLength = (int)numNoteTitleMaxLength.Value;
            _settings.NoteTitleIncludeDate = chkNoteTitleIncludeDate.Checked;

            // Save attachment settings
            _settings.AttachmentsFolder = txtAttachmentsFolder.Text.Trim();
            _settings.AttachmentStorageMode = (AttachmentStorageMode)cmbAttachmentStorageMode.SelectedIndex;
            _settings.SaveInlineImages = chkSaveInlineImages.Checked;
            _settings.SaveAllAttachments = chkSaveAllAttachments.Checked;
            _settings.UseObsidianWikilinks = chkUseObsidianWikilinks.Checked;

            _settings.Save();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (var form = new InputDialog("Add Pattern", "Enter regex pattern:"))
            {
                if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                {
                    lstPatterns.Items.Add(form.InputText);
                }
            }
        }

        private void BtnEdit_Click(object sender, EventArgs e)
        {
            if (lstPatterns.SelectedItem != null)
            {
                using (var form = new InputDialog("Edit Pattern", "Edit regex pattern:", lstPatterns.SelectedItem.ToString()))
                {
                    if (form.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(form.InputText))
                    {
                        int index = lstPatterns.SelectedIndex;
                        lstPatterns.Items[index] = form.InputText;
                    }
                }
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lstPatterns.SelectedItem != null)
            {
                lstPatterns.Items.RemoveAt(lstPatterns.SelectedIndex);
            }
        }

        private void chkEnableContactSaving_CheckedChanged(object sender, EventArgs e)
        {
            txtContactsFolder.Enabled = chkEnableContactSaving.Checked;
        }

        private void chkShowDevelopmentSettings_CheckedChanged(object sender, EventArgs e)
        {
            grpDevelopment.Visible = chkShowDevelopmentSettings.Checked;
            chkShowThreadDebug.Visible = chkShowDevelopmentSettings.Checked;
        }
    }
} 