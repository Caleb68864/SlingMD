using System;
using System.Windows.Forms;
using System.Drawing;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using System.Collections.Generic;
using System.Linq;

namespace SlingMD.Outlook.Forms
{
    public partial class SettingsForm : Form
    {
        private readonly ObsidianSettings _settings;

        // General tab controls
        private TextBox txtVaultName;
        private TextBox txtVaultPath;
        private Button btnBrowse;
        private CheckBox chkLaunchObsidian;
        private CheckBox chkShowCountdown;
        private NumericUpDown numDelay;
        private TextBox txtTemplatesFolder;
        private CheckBox chkIncludeDailyNoteLink;
        private TextBox txtDailyNoteLinkFormat;
        private Label lblDailyNoteLinkFormat;

        // Email tab controls
        private TextBox txtInboxFolder;
        private TextBox txtNoteTitleFormat;
        private NumericUpDown numNoteTitleMaxLength;
        private CheckBox chkNoteTitleIncludeDate;
        private Label lblNoteTitleFormat;
        private Label lblNoteTitleMaxLength;
        private Label lblNoteTitleIncludeDate;
        private TextBox txtDefaultNoteTags;
        private Label lblDefaultNoteTags;
        private ListBox lstPatterns;
        private Button btnAdd;
        private Button btnEdit;
        private Button btnRemove;
        private TextBox txtEmailFilenameFormat;
        private TextBox txtEmailTemplateFile;

        // Appointments tab controls
        private TextBox txtAppointmentsFolder;
        private TextBox txtAppointmentNoteTitleFormat;
        private NumericUpDown numAppointmentTitleMaxLength;
        private TextBox txtAppointmentDefaultTags;
        private CheckBox chkAppointmentSaveAttachments;
        private CheckBox chkCreateMeetingNotes;
        private TextBox txtMeetingNoteTemplate;
        private CheckBox chkGroupRecurringMeetings;
        private CheckBox chkSaveCancelledAppointments;
        private ComboBox cmbAppointmentTaskCreation;
        private TextBox txtAppointmentTemplateFile;

        // Contacts tab controls
        private TextBox txtContactsFolder;
        private CheckBox chkEnableContactSaving;
        private CheckBox chkSearchEntireVaultForContacts;
        private TextBox txtContactFilenameFormat;
        private TextBox txtContactTemplateFile;
        private CheckBox chkContactNoteIncludeDetails;
        private TextBox txtContactLinkFormat;
        private TextBox txtContactDateFormat;
        private TextBox txtEmailDateFormat;
        private TextBox txtAppointmentDateFormat;

        // Tasks tab controls
        private CheckBox chkCreateObsidianTask;
        private CheckBox chkCreateOutlookTask;
        private CheckBox chkAskForDates;
        private NumericUpDown numDefaultDueDays;
        private NumericUpDown numDefaultReminderDays;
        private NumericUpDown numDefaultReminderHour;
        private TextBox txtDefaultTaskTags;
        private Label lblDefaultTaskTags;
        private TextBox txtTaskTemplateFile;

        // Threading tab controls
        private CheckBox chkGroupEmailThreads;
        private CheckBox chkMoveDateToFrontInThread;
        private TextBox txtThreadTemplateFile;

        // Attachments tab controls
        private GroupBox grpAttachments;
        private TextBox txtAttachmentsFolder;
        private ComboBox cmbAttachmentStorageMode;
        private CheckBox chkSaveInlineImages;
        private CheckBox chkSaveAllAttachments;
        private CheckBox chkUseObsidianWikilinks;
        private Label lblAttachmentsFolder;
        private Label lblAttachmentStorageMode;

        // Developer tab controls
        private GroupBox grpDevelopment;
        private CheckBox chkShowDevelopmentSettings;
        private CheckBox chkShowThreadDebug;

        // Auto-Sling tab controls
        private CheckBox chkEnableAutoSling;
        private ComboBox cmbNotificationMode;
        private CheckBox chkEnableFlagToSling;
        private TextBox txtSentToObsidianCategory;
        private DataGridView dgvAutoSlingRules;
        private DataGridView dgvWatchedFolders;

        // Footer controls
        private Label lblSupportMessage;
        private LinkLabel lnkBuyMeACoffee;
        private Button btnSave;
        private Button btnCancel;

        // Layout containers
        private TableLayoutPanel rootLayout;
        private GroupBox grpNoteCustomization;
        private ToolTip toolTip;

        public SettingsForm(ObsidianSettings settings)
        {
            InitializeComponent();
            _settings = settings;
            LoadSettings();
        }

        /// <summary>
        /// Glyph appended to labels that have extended help available via tooltip + HelpForm.
        /// Circled Latin small letter i (U+24D8).
        /// </summary>
        private const string HelpIndicator = "  ⓘ";

        /// <summary>
        /// Binds a label (and optionally one or more controls) to a help entry. Appends the
        /// help indicator glyph to the label text and attaches a formatted tooltip to every
        /// provided control.
        /// </summary>
        private void BindHelp(string entryId, Label label, params Control[] controls)
        {
            HelpEntry entry = SettingsHelp.Get(entryId);
            if (entry == null) return;

            if (label != null && !label.Text.Contains("ⓘ"))
            {
                label.Text = label.Text.TrimEnd(':') + ":" + HelpIndicator;
            }

            string tip = SettingsHelp.FormatAsTooltip(entry);
            if (label != null) toolTip.SetToolTip(label, tip);
            if (controls != null)
            {
                foreach (Control c in controls)
                {
                    if (c != null) toolTip.SetToolTip(c, tip);
                }
            }
        }

        /// <summary>
        /// Binds a help entry to a control that has no separate label (e.g. a standalone
        /// CheckBox). The control's Text gets the help indicator appended.
        /// </summary>
        private void BindHelpInline(string entryId, Control control)
        {
            HelpEntry entry = SettingsHelp.Get(entryId);
            if (entry == null || control == null) return;

            if (!control.Text.Contains("ⓘ"))
            {
                control.Text = control.Text + HelpIndicator;
            }
            toolTip.SetToolTip(control, SettingsHelp.FormatAsTooltip(entry));
        }

        private void InitializeComponent()
        {
            this.toolTip = new ToolTip
            {
                AutoPopDelay = 30000,   // keep tooltip visible for 30s
                InitialDelay = 500,
                ReshowDelay = 200,
                ShowAlways = true
            };

            // Root layout: TabControl + Footer
            this.rootLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            this.rootLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            TabControl tabControl = new TabControl
            {
                Dock = DockStyle.Fill
            };

            // Create 9 tab pages
            TabPage tabGeneral = new TabPage("General");
            TabPage tabEmail = new TabPage("Email");
            TabPage tabAppointments = new TabPage("Appointments");
            TabPage tabContacts = new TabPage("Contacts");
            TabPage tabTasks = new TabPage("Tasks");
            TabPage tabThreading = new TabPage("Threading");
            TabPage tabAttachments = new TabPage("Attachments");
            TabPage tabAutoSling = new TabPage("Auto-Sling");
            TabPage tabDeveloper = new TabPage("Developer");

            tabControl.TabPages.AddRange(new TabPage[] {
                tabGeneral, tabEmail, tabAppointments, tabContacts,
                tabTasks, tabThreading, tabAttachments, tabAutoSling, tabDeveloper
            });

            // ---- General Tab ----
            TableLayoutPanel generalTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            generalTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            generalTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));

            int gRow = 0;
            Label lblVaultName = new Label { Text = "Vault Name:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(lblVaultName, 0, gRow);
            this.txtVaultName = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(this.txtVaultName, 1, gRow++);
            BindHelp("General.VaultName", lblVaultName, txtVaultName);

            Label lblVaultPath = new Label { Text = "Vault Base Path:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(lblVaultPath, 0, gRow);
            this.txtVaultPath = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(this.txtVaultPath, 1, gRow++);
            BindHelp("General.VaultBasePath", lblVaultPath, txtVaultPath);

            generalTabLayout.Controls.Add(new Label(), 0, gRow);
            this.btnBrowse = new Button { Text = "Browse...", Anchor = AnchorStyles.Left };
            this.btnBrowse.Click += btnBrowse_Click;
            generalTabLayout.Controls.Add(this.btnBrowse, 1, gRow++);

            this.chkLaunchObsidian = new CheckBox { Text = "Launch Obsidian after saving", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalTabLayout.Controls.Add(this.chkLaunchObsidian, 0, gRow);
            this.chkShowCountdown = new CheckBox { Text = "Show countdown", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            generalTabLayout.Controls.Add(this.chkShowCountdown, 1, gRow++);
            BindHelpInline("General.LaunchObsidian", chkLaunchObsidian);
            BindHelpInline("General.ShowCountdown", chkShowCountdown);

            Label lblDelay = new Label { Text = "Delay (seconds):", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(lblDelay, 0, gRow);
            this.numDelay = new NumericUpDown { Minimum = 0, Maximum = 10, Anchor = AnchorStyles.Left };
            generalTabLayout.Controls.Add(this.numDelay, 1, gRow++);
            BindHelp("General.Delay", lblDelay, numDelay);

            Label lblTemplatesFolder = new Label { Text = "Templates Folder:", Anchor = AnchorStyles.Left, AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(lblTemplatesFolder, 0, gRow);
            this.txtTemplatesFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(this.txtTemplatesFolder, 1, gRow++);
            BindHelp("General.TemplatesFolder", lblTemplatesFolder, txtTemplatesFolder);

            FlowLayoutPanel dailyLinkPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Anchor = AnchorStyles.Left };
            this.chkIncludeDailyNoteLink = new CheckBox { Text = "Include Daily Note Link", Anchor = AnchorStyles.Left, AutoSize = true };
            dailyLinkPanel.Controls.Add(this.chkIncludeDailyNoteLink);
            generalTabLayout.Controls.Add(new Label(), 0, gRow);
            generalTabLayout.Controls.Add(dailyLinkPanel, 1, gRow++);
            BindHelpInline("General.IncludeDailyNoteLink", chkIncludeDailyNoteLink);

            this.lblDailyNoteLinkFormat = new Label { Text = "Daily Note Link Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtDailyNoteLinkFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            generalTabLayout.Controls.Add(this.lblDailyNoteLinkFormat, 0, gRow);
            generalTabLayout.Controls.Add(this.txtDailyNoteLinkFormat, 1, gRow++);
            BindHelp("General.DailyNoteLinkFormat", lblDailyNoteLinkFormat, txtDailyNoteLinkFormat);

            this.chkIncludeDailyNoteLink.CheckedChanged += (s, e) =>
            {
                this.txtDailyNoteLinkFormat.Enabled = this.chkIncludeDailyNoteLink.Checked;
            };

            generalTabLayout.RowCount = gRow + 1;
            for (int i = 0; i < gRow; i++)
                generalTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            generalTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabGeneral.Controls.Add(generalTabLayout);

            // ---- Email Tab ----
            TableLayoutPanel emailTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            emailTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            emailTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));

            int eRow = 0;
            Label lblInboxFolder = new Label { Text = "Inbox Folder:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(lblInboxFolder, 0, eRow);
            this.txtInboxFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(this.txtInboxFolder, 1, eRow++);
            BindHelp("Email.InboxFolder", lblInboxFolder, txtInboxFolder);

            this.lblNoteTitleFormat = new Label { Text = "Note Title Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtNoteTitleFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(this.lblNoteTitleFormat, 0, eRow);
            emailTabLayout.Controls.Add(this.txtNoteTitleFormat, 1, eRow++);
            BindHelp("Email.NoteTitleFormat", lblNoteTitleFormat, txtNoteTitleFormat);

            this.lblNoteTitleMaxLength = new Label { Text = "Max Title Length:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.numNoteTitleMaxLength = new NumericUpDown { Minimum = 10, Maximum = 200, Anchor = AnchorStyles.Left };
            emailTabLayout.Controls.Add(this.lblNoteTitleMaxLength, 0, eRow);
            emailTabLayout.Controls.Add(this.numNoteTitleMaxLength, 1, eRow++);
            BindHelp("Email.NoteTitleMaxLength", lblNoteTitleMaxLength, numNoteTitleMaxLength);

            FlowLayoutPanel noteTitleDatePanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Anchor = AnchorStyles.Left };
            this.lblNoteTitleIncludeDate = new Label { Text = "Include Date in Title:", Anchor = AnchorStyles.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft };
            this.chkNoteTitleIncludeDate = new CheckBox { Anchor = AnchorStyles.Left };
            noteTitleDatePanel.Controls.Add(this.lblNoteTitleIncludeDate);
            noteTitleDatePanel.Controls.Add(this.chkNoteTitleIncludeDate);
            emailTabLayout.Controls.Add(new Label(), 0, eRow);
            emailTabLayout.Controls.Add(noteTitleDatePanel, 1, eRow++);
            BindHelp("Email.NoteTitleIncludeDate", lblNoteTitleIncludeDate, chkNoteTitleIncludeDate);

            this.lblDefaultNoteTags = new Label { Text = "Default Note Tags:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtDefaultNoteTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(this.lblDefaultNoteTags, 0, eRow);
            emailTabLayout.Controls.Add(this.txtDefaultNoteTags, 1, eRow++);
            BindHelp("Email.DefaultNoteTags", lblDefaultNoteTags, txtDefaultNoteTags);

            // Subject Cleanup Patterns section
            Label lblSubjectCleanupPatterns = new Label { Text = "Subject Cleanup Patterns:", Anchor = AnchorStyles.Left | AnchorStyles.Top, TextAlign = ContentAlignment.TopLeft, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(lblSubjectCleanupPatterns, 0, eRow);
            BindHelp("Email.SubjectCleanupPatterns", lblSubjectCleanupPatterns);
            TableLayoutPanel patternsPanel = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
            patternsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80F));
            patternsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20F));
            this.lstPatterns = new ListBox { Height = 120, Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom, Dock = DockStyle.Fill };
            patternsPanel.Controls.Add(this.lstPatterns, 0, 0);
            FlowLayoutPanel btnPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true };
            this.btnAdd = new Button { Text = "Add" };
            this.btnAdd.Click += BtnAdd_Click;
            this.btnEdit = new Button { Text = "Edit" };
            this.btnEdit.Click += BtnEdit_Click;
            this.btnRemove = new Button { Text = "Remove" };
            this.btnRemove.Click += BtnRemove_Click;
            btnPanel.Controls.Add(this.btnAdd);
            btnPanel.Controls.Add(this.btnEdit);
            btnPanel.Controls.Add(this.btnRemove);
            patternsPanel.Controls.Add(btnPanel, 1, 0);
            emailTabLayout.Controls.Add(patternsPanel, 1, eRow++);

            Label lblEmailFilenameFormat = new Label { Text = "Email Filename Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(lblEmailFilenameFormat, 0, eRow);
            this.txtEmailFilenameFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(this.txtEmailFilenameFormat, 1, eRow++);
            BindHelp("Email.EmailFilenameFormat", lblEmailFilenameFormat, txtEmailFilenameFormat);

            Label lblEmailTemplateFile = new Label { Text = "Email Template File:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(lblEmailTemplateFile, 0, eRow);
            this.txtEmailTemplateFile = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(this.txtEmailTemplateFile, 1, eRow++);
            BindHelp("Email.EmailTemplateFile", lblEmailTemplateFile, txtEmailTemplateFile);

            Label lblEmailDateFormat = new Label { Text = "Email Date Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(lblEmailDateFormat, 0, eRow);
            this.txtEmailDateFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            emailTabLayout.Controls.Add(this.txtEmailDateFormat, 1, eRow++);
            BindHelp("Email.EmailDateFormat", lblEmailDateFormat, txtEmailDateFormat);

            emailTabLayout.RowCount = eRow + 1;
            for (int i = 0; i < eRow; i++)
                emailTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            emailTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabEmail.Controls.Add(emailTabLayout);

            // ---- Appointments Tab ----
            TableLayoutPanel apptTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            apptTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            apptTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));

            int aRow = 0;
            Label lblApptFolder = new Label { Text = "Appointments Folder:", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptFolder, 0, aRow);
            this.txtAppointmentsFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(this.txtAppointmentsFolder, 1, aRow++);
            BindHelp("Appointments.AppointmentsFolder", lblApptFolder, txtAppointmentsFolder);

            Label lblApptTitleFormat = new Label { Text = "Note Title Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptTitleFormat, 0, aRow);
            this.txtAppointmentNoteTitleFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(this.txtAppointmentNoteTitleFormat, 1, aRow++);
            BindHelp("Appointments.AppointmentNoteTitleFormat", lblApptTitleFormat, txtAppointmentNoteTitleFormat);

            Label lblApptMaxLength = new Label { Text = "Max Title Length:", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptMaxLength, 0, aRow);
            this.numAppointmentTitleMaxLength = new NumericUpDown { Minimum = 10, Maximum = 500, Value = 50, Anchor = AnchorStyles.Left };
            apptTabLayout.Controls.Add(this.numAppointmentTitleMaxLength, 1, aRow++);
            BindHelp("Appointments.AppointmentNoteTitleMaxLength", lblApptMaxLength, numAppointmentTitleMaxLength);

            Label lblApptDefaultTags = new Label { Text = "Default Tags:", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptDefaultTags, 0, aRow);
            this.txtAppointmentDefaultTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(this.txtAppointmentDefaultTags, 1, aRow++);
            BindHelp("Appointments.AppointmentDefaultTags", lblApptDefaultTags, txtAppointmentDefaultTags);

            Label lblApptTemplate = new Label { Text = "Appointment Template File:", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptTemplate, 0, aRow);
            this.txtAppointmentTemplateFile = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(this.txtAppointmentTemplateFile, 1, aRow++);
            BindHelp("Appointments.AppointmentTemplateFile", lblApptTemplate, txtAppointmentTemplateFile);

            Label lblMeetingNoteTemplate = new Label { Text = "Meeting Note Template:", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblMeetingNoteTemplate, 0, aRow);
            this.txtMeetingNoteTemplate = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(this.txtMeetingNoteTemplate, 1, aRow++);
            BindHelp("Appointments.MeetingNoteTemplate", lblMeetingNoteTemplate, txtMeetingNoteTemplate);

            Label lblApptDateFormat = new Label { Text = "Appointment Date Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptDateFormat, 0, aRow);
            this.txtAppointmentDateFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(this.txtAppointmentDateFormat, 1, aRow++);
            BindHelp("Appointments.AppointmentDateFormat", lblApptDateFormat, txtAppointmentDateFormat);

            Label lblApptTaskCreation = new Label { Text = "Task Creation Mode:", AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            apptTabLayout.Controls.Add(lblApptTaskCreation, 0, aRow);
            this.cmbAppointmentTaskCreation = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right
            };
            this.cmbAppointmentTaskCreation.Items.AddRange(new object[] { "None", "Obsidian", "Outlook", "Both" });
            apptTabLayout.Controls.Add(this.cmbAppointmentTaskCreation, 1, aRow++);
            BindHelp("Appointments.AppointmentTaskCreation", lblApptTaskCreation, cmbAppointmentTaskCreation);

            apptTabLayout.Controls.Add(new Label(), 0, aRow);
            FlowLayoutPanel apptCheckboxPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, Anchor = AnchorStyles.Left };
            this.chkAppointmentSaveAttachments = new CheckBox { Text = "Save attachments", AutoSize = true };
            this.chkCreateMeetingNotes = new CheckBox { Text = "Create companion meeting notes", AutoSize = true };
            this.chkGroupRecurringMeetings = new CheckBox { Text = "Group recurring meeting instances", AutoSize = true };
            this.chkSaveCancelledAppointments = new CheckBox { Text = "Save cancelled appointments", AutoSize = true };
            apptCheckboxPanel.Controls.Add(this.chkAppointmentSaveAttachments);
            apptCheckboxPanel.Controls.Add(this.chkCreateMeetingNotes);
            apptCheckboxPanel.Controls.Add(this.chkGroupRecurringMeetings);
            apptCheckboxPanel.Controls.Add(this.chkSaveCancelledAppointments);
            apptTabLayout.Controls.Add(apptCheckboxPanel, 1, aRow++);
            BindHelpInline("Appointments.AppointmentSaveAttachments", chkAppointmentSaveAttachments);
            BindHelpInline("Appointments.CreateMeetingNotes", chkCreateMeetingNotes);
            BindHelpInline("Appointments.GroupRecurringMeetings", chkGroupRecurringMeetings);
            BindHelpInline("Appointments.SaveCancelledAppointments", chkSaveCancelledAppointments);

            apptTabLayout.RowCount = aRow + 1;
            for (int i = 0; i < aRow; i++)
                apptTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            apptTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabAppointments.Controls.Add(apptTabLayout);

            // ---- Contacts Tab ----
            TableLayoutPanel contactsTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            contactsTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            contactsTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));

            int cRow = 0;
            Label lblContactsFolder = new Label { Text = "Contacts Folder:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(lblContactsFolder, 0, cRow);
            this.txtContactsFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(this.txtContactsFolder, 1, cRow++);
            BindHelp("Contacts.ContactsFolder", lblContactsFolder, txtContactsFolder);

            contactsTabLayout.Controls.Add(new Label(), 0, cRow);
            this.chkEnableContactSaving = new CheckBox { Text = "Enable Contact Saving", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            this.chkEnableContactSaving.CheckedChanged += chkEnableContactSaving_CheckedChanged;
            contactsTabLayout.Controls.Add(this.chkEnableContactSaving, 1, cRow++);
            BindHelpInline("Contacts.EnableContactSaving", chkEnableContactSaving);

            contactsTabLayout.Controls.Add(new Label(), 0, cRow);
            this.chkSearchEntireVaultForContacts = new CheckBox { Text = "Search entire vault for contacts", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            contactsTabLayout.Controls.Add(this.chkSearchEntireVaultForContacts, 1, cRow++);
            BindHelpInline("Contacts.SearchEntireVaultForContacts", chkSearchEntireVaultForContacts);

            Label lblContactFilenameFormat = new Label { Text = "Contact Filename Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(lblContactFilenameFormat, 0, cRow);
            this.txtContactFilenameFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(this.txtContactFilenameFormat, 1, cRow++);
            BindHelp("Contacts.ContactFilenameFormat", lblContactFilenameFormat, txtContactFilenameFormat);

            Label lblContactTemplateFile = new Label { Text = "Contact Template File:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(lblContactTemplateFile, 0, cRow);
            this.txtContactTemplateFile = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(this.txtContactTemplateFile, 1, cRow++);
            BindHelp("Contacts.ContactTemplateFile", lblContactTemplateFile, txtContactTemplateFile);

            contactsTabLayout.Controls.Add(new Label(), 0, cRow);
            this.chkContactNoteIncludeDetails = new CheckBox { Text = "Include contact details (phone, email, company, etc.)", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            contactsTabLayout.Controls.Add(this.chkContactNoteIncludeDetails, 1, cRow++);
            BindHelpInline("Contacts.ContactNoteIncludeDetails", chkContactNoteIncludeDetails);

            Label lblContactLinkFormat = new Label { Text = "Contact Link Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(lblContactLinkFormat, 0, cRow);
            this.txtContactLinkFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(this.txtContactLinkFormat, 1, cRow++);
            BindHelp("Contacts.ContactLinkFormat", lblContactLinkFormat, txtContactLinkFormat);

            Label lblContactDateFormat = new Label { Text = "Contact Date Format:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(lblContactDateFormat, 0, cRow);
            this.txtContactDateFormat = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            contactsTabLayout.Controls.Add(this.txtContactDateFormat, 1, cRow++);
            BindHelp("Contacts.ContactDateFormat", lblContactDateFormat, txtContactDateFormat);

            contactsTabLayout.RowCount = cRow + 1;
            for (int i = 0; i < cRow; i++)
                contactsTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            contactsTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabContacts.Controls.Add(contactsTabLayout);

            // ---- Tasks Tab ----
            TableLayoutPanel tasksTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            tasksTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            tasksTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));

            int tRow = 0;
            tasksTabLayout.Controls.Add(new Label(), 0, tRow);
            FlowLayoutPanel taskCheckPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Anchor = AnchorStyles.Left };
            this.chkCreateObsidianTask = new CheckBox { Text = "Create task in Obsidian note", AutoSize = true };
            this.chkCreateOutlookTask = new CheckBox { Text = "Create task in Outlook", AutoSize = true, Margin = new Padding(12, 0, 0, 0) };
            taskCheckPanel.Controls.Add(this.chkCreateObsidianTask);
            taskCheckPanel.Controls.Add(this.chkCreateOutlookTask);
            tasksTabLayout.Controls.Add(taskCheckPanel, 1, tRow++);
            BindHelpInline("Tasks.CreateObsidianTask", chkCreateObsidianTask);
            BindHelpInline("Tasks.CreateOutlookTask", chkCreateOutlookTask);

            tasksTabLayout.Controls.Add(new Label(), 0, tRow);
            this.chkAskForDates = new CheckBox { Text = "Ask for dates and times each time", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            tasksTabLayout.Controls.Add(this.chkAskForDates, 1, tRow++);
            BindHelpInline("Tasks.AskForDates", chkAskForDates);

            Label lblDueInDays = new Label { Text = "Due in Days:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            tasksTabLayout.Controls.Add(lblDueInDays, 0, tRow);
            this.numDefaultDueDays = new NumericUpDown { Minimum = 0, Maximum = 30, Anchor = AnchorStyles.Left };
            tasksTabLayout.Controls.Add(this.numDefaultDueDays, 1, tRow++);
            BindHelp("Tasks.DueInDays", lblDueInDays, numDefaultDueDays);

            Label lblReminderDays = new Label { Text = "Reminder Days:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            tasksTabLayout.Controls.Add(lblReminderDays, 0, tRow);
            this.numDefaultReminderDays = new NumericUpDown { Minimum = 0, Maximum = 30, Anchor = AnchorStyles.Left };
            tasksTabLayout.Controls.Add(this.numDefaultReminderDays, 1, tRow++);
            BindHelp("Tasks.ReminderDays", lblReminderDays, numDefaultReminderDays);

            Label lblReminderHour = new Label { Text = "Reminder Hour:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            tasksTabLayout.Controls.Add(lblReminderHour, 0, tRow);
            this.numDefaultReminderHour = new NumericUpDown { Minimum = 0, Maximum = 23, Anchor = AnchorStyles.Left };
            tasksTabLayout.Controls.Add(this.numDefaultReminderHour, 1, tRow++);
            BindHelp("Tasks.ReminderHour", lblReminderHour, numDefaultReminderHour);

            this.lblDefaultTaskTags = new Label { Text = "Default Task Tags:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            this.txtDefaultTaskTags = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            tasksTabLayout.Controls.Add(this.lblDefaultTaskTags, 0, tRow);
            tasksTabLayout.Controls.Add(this.txtDefaultTaskTags, 1, tRow++);
            BindHelp("Tasks.DefaultTaskTags", lblDefaultTaskTags, txtDefaultTaskTags);

            Label lblTaskTemplate = new Label { Text = "Task Template File:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            tasksTabLayout.Controls.Add(lblTaskTemplate, 0, tRow);
            this.txtTaskTemplateFile = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            tasksTabLayout.Controls.Add(this.txtTaskTemplateFile, 1, tRow++);
            BindHelp("Tasks.TaskTemplateFile", lblTaskTemplate, txtTaskTemplateFile);

            tasksTabLayout.RowCount = tRow + 1;
            for (int i = 0; i < tRow; i++)
                tasksTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tasksTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabTasks.Controls.Add(tasksTabLayout);

            // ---- Threading Tab ----
            TableLayoutPanel threadingTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            threadingTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40F));
            threadingTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60F));

            int thRow = 0;
            threadingTabLayout.Controls.Add(new Label(), 0, thRow);
            this.chkGroupEmailThreads = new CheckBox { Text = "Group email threads", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            threadingTabLayout.Controls.Add(this.chkGroupEmailThreads, 1, thRow++);
            BindHelpInline("Threading.GroupEmailThreads", chkGroupEmailThreads);

            threadingTabLayout.Controls.Add(new Label(), 0, thRow);
            this.chkMoveDateToFrontInThread = new CheckBox { Text = "Move date to front of filename when grouping threads", Anchor = AnchorStyles.Left, AutoSize = true };
            threadingTabLayout.Controls.Add(this.chkMoveDateToFrontInThread, 1, thRow++);
            BindHelpInline("Threading.MoveDateToFrontInThread", chkMoveDateToFrontInThread);

            Label lblThreadTemplate = new Label { Text = "Thread Template File:", AutoSize = false, AutoEllipsis = true, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            threadingTabLayout.Controls.Add(lblThreadTemplate, 0, thRow);
            this.txtThreadTemplateFile = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            threadingTabLayout.Controls.Add(this.txtThreadTemplateFile, 1, thRow++);
            BindHelp("Threading.ThreadTemplateFile", lblThreadTemplate, txtThreadTemplateFile);

            // Add event handler for enabling/disabling move date checkbox
            this.chkNoteTitleIncludeDate.CheckedChanged += (s, e) =>
            {
                if (!chkNoteTitleIncludeDate.Checked)
                {
                    chkMoveDateToFrontInThread.Checked = false;
                    chkMoveDateToFrontInThread.Enabled = false;
                }
                else
                {
                    chkMoveDateToFrontInThread.Enabled = true;
                }
            };

            threadingTabLayout.RowCount = thRow + 1;
            for (int i = 0; i < thRow; i++)
                threadingTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            threadingTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabThreading.Controls.Add(threadingTabLayout);

            // ---- Attachments Tab ----
            this.grpAttachments = new GroupBox
            {
                Text = "Attachment Settings",
                Dock = DockStyle.Top,
                AutoSize = true
            };
            TableLayoutPanel attachmentLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, AutoSize = true };
            attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            attachmentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));

            this.lblAttachmentStorageMode = new Label { Text = "Storage Location:", Anchor = AnchorStyles.Left };
            this.cmbAttachmentStorageMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Width = 320
            };
            this.cmbAttachmentStorageMode.Items.Add("Same folder as note");
            this.cmbAttachmentStorageMode.Items.Add("Subfolder per note");
            this.cmbAttachmentStorageMode.Items.Add("Centralized folder");
            attachmentLayout.Controls.Add(this.lblAttachmentStorageMode, 0, 0);
            attachmentLayout.Controls.Add(this.cmbAttachmentStorageMode, 1, 0);
            BindHelp("Attachments.StorageMode", lblAttachmentStorageMode, cmbAttachmentStorageMode);

            this.lblAttachmentsFolder = new Label { Text = "Centralized Folder Name:", Anchor = AnchorStyles.Left };
            this.txtAttachmentsFolder = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Width = 320 };
            attachmentLayout.Controls.Add(this.lblAttachmentsFolder, 0, 1);
            attachmentLayout.Controls.Add(this.txtAttachmentsFolder, 1, 1);
            BindHelp("Attachments.AttachmentsFolder", lblAttachmentsFolder, txtAttachmentsFolder);

            FlowLayoutPanel checkboxLayout = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Left
            };
            this.chkSaveInlineImages = new CheckBox { Text = "Save inline images", AutoSize = true };
            this.chkSaveAllAttachments = new CheckBox { Text = "Save all attachments", AutoSize = true, Margin = new Padding(20, 0, 0, 0) };
            this.chkUseObsidianWikilinks = new CheckBox { Text = "Use Obsidian wikilinks", AutoSize = true, Margin = new Padding(20, 0, 0, 0) };
            checkboxLayout.Controls.Add(this.chkSaveInlineImages);
            checkboxLayout.Controls.Add(this.chkSaveAllAttachments);
            checkboxLayout.Controls.Add(this.chkUseObsidianWikilinks);
            attachmentLayout.Controls.Add(checkboxLayout, 1, 2);
            BindHelpInline("Attachments.SaveInlineImages", chkSaveInlineImages);
            BindHelpInline("Attachments.SaveAllAttachments", chkSaveAllAttachments);
            BindHelpInline("Attachments.UseObsidianWikilinks", chkUseObsidianWikilinks);

            this.cmbAttachmentStorageMode.SelectedIndexChanged += (s, e) =>
            {
                this.txtAttachmentsFolder.Enabled = this.cmbAttachmentStorageMode.SelectedIndex == 2;
            };

            this.grpAttachments.Controls.Add(attachmentLayout);

            TableLayoutPanel attachmentsTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            attachmentsTabLayout.Controls.Add(this.grpAttachments);
            tabAttachments.Controls.Add(attachmentsTabLayout);

            // ---- Auto-Sling Tab ----
            TableLayoutPanel autoSlingTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            autoSlingTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35F));
            autoSlingTabLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 65F));

            int asRow = 0;

            autoSlingTabLayout.Controls.Add(new Label(), 0, asRow);
            this.chkEnableAutoSling = new CheckBox { Text = "Enable Auto-Sling", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            autoSlingTabLayout.Controls.Add(this.chkEnableAutoSling, 1, asRow++);
            BindHelpInline("AutoSling.EnableAutoSling", chkEnableAutoSling);

            Label lblNotifMode = new Label { Text = "Notification Mode:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            autoSlingTabLayout.Controls.Add(lblNotifMode, 0, asRow);
            this.cmbNotificationMode = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Dock = DockStyle.Fill
            };
            this.cmbNotificationMode.Items.Add("Toast");
            this.cmbNotificationMode.Items.Add("Silent");
            autoSlingTabLayout.Controls.Add(this.cmbNotificationMode, 1, asRow++);
            BindHelp("AutoSling.NotificationMode", lblNotifMode, cmbNotificationMode);

            autoSlingTabLayout.Controls.Add(new Label(), 0, asRow);
            this.chkEnableFlagToSling = new CheckBox { Text = "Enable Flag-to-Sling (auto-sling flagged emails)", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            autoSlingTabLayout.Controls.Add(this.chkEnableFlagToSling, 1, asRow++);
            BindHelpInline("AutoSling.EnableFlagToSling", chkEnableFlagToSling);

            Label lblSentToObsidianCat = new Label { Text = "\"Sent to Obsidian\" Category:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            autoSlingTabLayout.Controls.Add(lblSentToObsidianCat, 0, asRow);
            this.txtSentToObsidianCategory = new TextBox { Anchor = AnchorStyles.Left | AnchorStyles.Right, Dock = DockStyle.Fill };
            autoSlingTabLayout.Controls.Add(this.txtSentToObsidianCategory, 1, asRow++);
            BindHelp("AutoSling.SentToObsidianCategory", lblSentToObsidianCat, txtSentToObsidianCategory);

            Label lblAutoSlingRules = new Label { Text = "Auto-Sling Rules:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            autoSlingTabLayout.Controls.Add(lblAutoSlingRules, 0, asRow);
            autoSlingTabLayout.Controls.Add(new Label(), 1, asRow++);
            BindHelp("AutoSling.Rules", lblAutoSlingRules);

            this.dgvAutoSlingRules = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                Height = 120
            };
            DataGridViewComboBoxColumn ruleTypeCol = new DataGridViewComboBoxColumn
            {
                HeaderText = "Type",
                Name = "Type",
                FillWeight = 30
            };
            ruleTypeCol.Items.AddRange("Sender", "Domain", "Category");
            DataGridViewTextBoxColumn rulePatternCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Pattern",
                Name = "Pattern",
                FillWeight = 55
            };
            DataGridViewCheckBoxColumn ruleEnabledCol = new DataGridViewCheckBoxColumn
            {
                HeaderText = "Enabled",
                Name = "Enabled",
                FillWeight = 15
            };
            this.dgvAutoSlingRules.Columns.Add(ruleTypeCol);
            this.dgvAutoSlingRules.Columns.Add(rulePatternCol);
            this.dgvAutoSlingRules.Columns.Add(ruleEnabledCol);
            autoSlingTabLayout.SetColumnSpan(this.dgvAutoSlingRules, 2);
            autoSlingTabLayout.Controls.Add(this.dgvAutoSlingRules, 0, asRow++);

            Label lblWatchedFolders = new Label { Text = "Watched Folders:", Anchor = AnchorStyles.Left, TextAlign = ContentAlignment.MiddleLeft, Dock = DockStyle.Fill };
            autoSlingTabLayout.Controls.Add(lblWatchedFolders, 0, asRow);
            autoSlingTabLayout.Controls.Add(new Label(), 1, asRow++);
            BindHelp("AutoSling.WatchedFolders", lblWatchedFolders);

            this.dgvWatchedFolders = new DataGridView
            {
                Dock = DockStyle.Fill,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                AllowUserToAddRows = true,
                AllowUserToDeleteRows = true,
                Height = 120
            };
            DataGridViewTextBoxColumn folderPathCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Folder Path",
                Name = "FolderPath",
                FillWeight = 50
            };
            DataGridViewTextBoxColumn folderTemplateCol = new DataGridViewTextBoxColumn
            {
                HeaderText = "Custom Template",
                Name = "CustomTemplate",
                FillWeight = 35
            };
            DataGridViewCheckBoxColumn folderEnabledCol = new DataGridViewCheckBoxColumn
            {
                HeaderText = "Enabled",
                Name = "Enabled",
                FillWeight = 15
            };
            this.dgvWatchedFolders.Columns.Add(folderPathCol);
            this.dgvWatchedFolders.Columns.Add(folderTemplateCol);
            this.dgvWatchedFolders.Columns.Add(folderEnabledCol);
            autoSlingTabLayout.SetColumnSpan(this.dgvWatchedFolders, 2);
            autoSlingTabLayout.Controls.Add(this.dgvWatchedFolders, 0, asRow++);

            autoSlingTabLayout.RowCount = asRow + 1;
            for (int i = 0; i < asRow; i++)
                autoSlingTabLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            autoSlingTabLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            tabAutoSling.Controls.Add(autoSlingTabLayout);

            // ---- Developer Tab ----
            this.grpDevelopment = new GroupBox
            {
                Text = "Development Settings",
                Dock = DockStyle.Top,
                AutoSize = true
            };
            FlowLayoutPanel devLayout = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true };
            this.chkShowDevelopmentSettings = new CheckBox { Text = "Show development settings", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            this.chkShowDevelopmentSettings.CheckedChanged += chkShowDevelopmentSettings_CheckedChanged;
            this.chkShowThreadDebug = new CheckBox { Text = "Show thread debug", Anchor = AnchorStyles.Left | AnchorStyles.Right, AutoSize = true };
            devLayout.Controls.Add(this.chkShowDevelopmentSettings);
            devLayout.Controls.Add(this.chkShowThreadDebug);
            BindHelpInline("Developer.ShowDevelopmentSettings", chkShowDevelopmentSettings);
            BindHelpInline("Developer.ShowThreadDebug", chkShowThreadDebug);
            this.grpDevelopment.Controls.Add(devLayout);

            TableLayoutPanel developerTabLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                AutoSize = true,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            developerTabLayout.Controls.Add(this.grpDevelopment);
            tabDeveloper.Controls.Add(developerTabLayout);

            // ---- Footer ----
            TableLayoutPanel footerLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                ColumnCount = 2
            };
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            footerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footerLayout.Padding = new Padding(12, 10, 12, 10);

            FlowLayoutPanel supportLayout = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };
            this.lblSupportMessage = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(460, 0),
                Text = SupportService.GetSupportMessage(),
                Margin = new Padding(0, 0, 0, 2)
            };
            this.lnkBuyMeACoffee = new LinkLabel
            {
                AutoSize = true,
                Text = SupportService.BuyMeACoffeeUrl,
                LinkBehavior = LinkBehavior.HoverUnderline,
                Margin = new Padding(0)
            };
            this.lnkBuyMeACoffee.LinkClicked += lnkBuyMeACoffee_LinkClicked;
            supportLayout.Controls.Add(this.lblSupportMessage);
            supportLayout.Controls.Add(this.lnkBuyMeACoffee);
            footerLayout.Controls.Add(supportLayout, 0, 0);

            FlowLayoutPanel btnLayout = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Margin = new Padding(12, 0, 0, 0)
            };
            this.btnSave = new Button { Text = "Save", DialogResult = DialogResult.OK };
            this.btnSave.Click += btnSave_Click;
            this.btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
            Button btnHelp = new Button { Text = "Help", Margin = new Padding(0, 0, 8, 0) };
            btnHelp.Click += (s, e) => { new HelpForm().Show(this); };
            toolTip.SetToolTip(btnHelp, "Open the searchable settings reference (Ctrl+F to focus the search box).");
            btnLayout.Controls.Add(this.btnSave);
            btnLayout.Controls.Add(this.btnCancel);
            btnLayout.Controls.Add(btnHelp);
            footerLayout.Controls.Add(btnLayout, 1, 0);

            // Assemble root
            this.rootLayout.Controls.Add(tabControl, 0, 0);
            this.rootLayout.Controls.Add(footerLayout, 0, 1);

            // Set up the form
            this.Controls.Add(this.rootLayout);
            this.AcceptButton = this.btnSave;
            this.CancelButton = this.btnCancel;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MaximizeBox = true;
            this.MinimizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Text = "Obsidian Settings";
            this.ClientSize = new Size(760, 820);

            System.IO.Stream iconStream = System.Reflection.Assembly.GetExecutingAssembly()
                .GetManifestResourceStream("SlingMD.Outlook.Resources.SlingMD.ico");
            if (iconStream != null)
            {
                this.Icon = new Icon(iconStream);
            }

            // grpNoteCustomization is kept as a field reference (unused in tabs, legacy field kept for compatibility)
            this.grpNoteCustomization = new GroupBox();
        }

        private void LoadSettings()
        {
            // General tab
            txtVaultName.Text = _settings.VaultName;
            txtVaultPath.Text = _settings.VaultBasePath;
            chkLaunchObsidian.Checked = _settings.LaunchObsidian;
            chkShowCountdown.Checked = _settings.ShowCountdown;
            numDelay.Value = _settings.ObsidianDelaySeconds;
            txtTemplatesFolder.Text = _settings.TemplatesFolder ?? "Templates";
            chkIncludeDailyNoteLink.Checked = _settings.IncludeDailyNoteLink;
            txtDailyNoteLinkFormat.Text = _settings.DailyNoteLinkFormat ?? "[[yyyy-MM-dd]]";
            txtDailyNoteLinkFormat.Enabled = _settings.IncludeDailyNoteLink;

            // Email tab
            txtInboxFolder.Text = _settings.InboxFolder;
            txtNoteTitleFormat.Text = _settings.NoteTitleFormat ?? "{Subject} - {Date}";
            numNoteTitleMaxLength.Value = _settings.NoteTitleMaxLength > 0 ? _settings.NoteTitleMaxLength : 50;
            chkNoteTitleIncludeDate.Checked = _settings.NoteTitleIncludeDate;
            txtDefaultNoteTags.Text = string.Join(", ", _settings.DefaultNoteTags ?? new List<string>());
            lstPatterns.Items.Clear();
            foreach (string pattern in _settings.SubjectCleanupPatterns)
            {
                lstPatterns.Items.Add(pattern);
            }
            txtEmailFilenameFormat.Text = _settings.EmailFilenameFormat ?? string.Empty;
            txtEmailTemplateFile.Text = _settings.EmailTemplateFile ?? "EmailTemplate.md";

            // Appointments tab
            txtAppointmentsFolder.Text = _settings.AppointmentsFolder;
            txtAppointmentNoteTitleFormat.Text = _settings.AppointmentNoteTitleFormat;
            numAppointmentTitleMaxLength.Value = _settings.AppointmentNoteTitleMaxLength > 0 ? _settings.AppointmentNoteTitleMaxLength : 50;
            txtAppointmentDefaultTags.Text = string.Join(", ", _settings.AppointmentDefaultNoteTags ?? new List<string>());
            chkAppointmentSaveAttachments.Checked = _settings.AppointmentSaveAttachments;
            chkCreateMeetingNotes.Checked = _settings.CreateMeetingNotes;
            txtMeetingNoteTemplate.Text = _settings.MeetingNoteTemplate ?? string.Empty;
            chkGroupRecurringMeetings.Checked = _settings.GroupRecurringMeetings;
            chkSaveCancelledAppointments.Checked = _settings.SaveCancelledAppointments;
            string apptTaskCreation = _settings.AppointmentTaskCreation ?? "None";
            int apptTaskIdx = cmbAppointmentTaskCreation.Items.IndexOf(apptTaskCreation);
            cmbAppointmentTaskCreation.SelectedIndex = apptTaskIdx >= 0 ? apptTaskIdx : 0;
            txtAppointmentTemplateFile.Text = _settings.AppointmentTemplateFile ?? "AppointmentTemplate.md";

            // Contacts tab
            txtContactsFolder.Text = _settings.ContactsFolder;
            chkEnableContactSaving.Checked = _settings.EnableContactSaving;
            txtContactsFolder.Enabled = _settings.EnableContactSaving;
            chkSearchEntireVaultForContacts.Checked = _settings.SearchEntireVaultForContacts;
            txtContactFilenameFormat.Text = _settings.ContactFilenameFormat ?? "{ContactName}";
            txtContactTemplateFile.Text = _settings.ContactTemplateFile ?? "ContactTemplate.md";
            chkContactNoteIncludeDetails.Checked = _settings.ContactNoteIncludeDetails;
            txtContactLinkFormat.Text = _settings.ContactLinkFormat ?? "[[{FullName}]]";
            txtContactDateFormat.Text = _settings.ContactDateFormat ?? "yyyy-MM-dd";
            txtEmailDateFormat.Text = _settings.EmailDateFormat ?? "yyyy-MM-dd HH:mm:ss";
            txtAppointmentDateFormat.Text = _settings.AppointmentDateFormat ?? "yyyy-MM-dd HH:mm";

            // Tasks tab
            chkCreateObsidianTask.Checked = _settings.CreateObsidianTask;
            chkCreateOutlookTask.Checked = _settings.CreateOutlookTask;
            chkAskForDates.Checked = _settings.AskForDates;
            numDefaultDueDays.Value = _settings.DefaultDueDays;
            numDefaultReminderDays.Value = _settings.DefaultReminderDays;
            numDefaultReminderHour.Value = _settings.DefaultReminderHour;
            txtDefaultTaskTags.Text = string.Join(", ", _settings.DefaultTaskTags ?? new List<string>());
            txtTaskTemplateFile.Text = _settings.TaskTemplateFile ?? "TaskTemplate.md";

            // Threading tab
            chkGroupEmailThreads.Checked = _settings.GroupEmailThreads;
            chkMoveDateToFrontInThread.Checked = _settings.MoveDateToFrontInThread;
            chkMoveDateToFrontInThread.Enabled = _settings.NoteTitleIncludeDate;
            txtThreadTemplateFile.Text = _settings.ThreadTemplateFile ?? "ThreadNoteTemplate.md";

            // Attachments tab
            txtAttachmentsFolder.Text = _settings.AttachmentsFolder ?? "Attachments";
            cmbAttachmentStorageMode.SelectedIndex = (int)_settings.AttachmentStorageMode;
            txtAttachmentsFolder.Enabled = _settings.AttachmentStorageMode == AttachmentStorageMode.Centralized;
            chkSaveInlineImages.Checked = _settings.SaveInlineImages;
            chkSaveAllAttachments.Checked = _settings.SaveAllAttachments;
            chkUseObsidianWikilinks.Checked = _settings.UseObsidianWikilinks;

            // Developer tab
            chkShowDevelopmentSettings.Checked = _settings.ShowDevelopmentSettings;
            chkShowThreadDebug.Checked = _settings.ShowThreadDebug;
            grpDevelopment.Visible = true;
            chkShowThreadDebug.Visible = _settings.ShowDevelopmentSettings;

            // Auto-Sling tab
            chkEnableAutoSling.Checked = _settings.EnableAutoSling;
            int notifIdx = cmbNotificationMode.Items.IndexOf(_settings.AutoSlingNotificationMode ?? "Toast");
            cmbNotificationMode.SelectedIndex = notifIdx >= 0 ? notifIdx : 0;
            chkEnableFlagToSling.Checked = _settings.EnableFlagToSling;
            txtSentToObsidianCategory.Text = _settings.SentToObsidianCategory ?? "Sent to Obsidian";

            dgvAutoSlingRules.Rows.Clear();
            foreach (AutoSlingRule rule in _settings.AutoSlingRules ?? new List<AutoSlingRule>())
            {
                dgvAutoSlingRules.Rows.Add(rule.Type, rule.Pattern, rule.Enabled);
            }

            dgvWatchedFolders.Rows.Clear();
            foreach (WatchedFolder folder in _settings.WatchedFolders ?? new List<WatchedFolder>())
            {
                dgvWatchedFolders.Rows.Add(folder.FolderPath, folder.CustomTemplate, folder.Enabled);
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
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
            // General tab
            _settings.VaultName = txtVaultName.Text;
            _settings.VaultBasePath = txtVaultPath.Text;
            _settings.LaunchObsidian = chkLaunchObsidian.Checked;
            _settings.ShowCountdown = chkShowCountdown.Checked;
            _settings.ObsidianDelaySeconds = (int)numDelay.Value;
            _settings.TemplatesFolder = txtTemplatesFolder.Text.Trim();
            _settings.IncludeDailyNoteLink = chkIncludeDailyNoteLink.Checked;
            _settings.DailyNoteLinkFormat = txtDailyNoteLinkFormat.Text.Trim();

            // Email tab
            _settings.InboxFolder = txtInboxFolder.Text;
            _settings.NoteTitleFormat = txtNoteTitleFormat.Text.Trim();
            _settings.NoteTitleMaxLength = (int)numNoteTitleMaxLength.Value;
            _settings.NoteTitleIncludeDate = chkNoteTitleIncludeDate.Checked;
            _settings.DefaultNoteTags = txtDefaultNoteTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)).ToList();
            _settings.SubjectCleanupPatterns.Clear();
            foreach (string pattern in lstPatterns.Items)
            {
                _settings.SubjectCleanupPatterns.Add(pattern);
            }
            _settings.EmailFilenameFormat = txtEmailFilenameFormat.Text.Trim();
            _settings.EmailTemplateFile = txtEmailTemplateFile.Text.Trim();

            // Appointments tab
            _settings.AppointmentsFolder = txtAppointmentsFolder.Text;
            _settings.AppointmentNoteTitleFormat = txtAppointmentNoteTitleFormat.Text;
            _settings.AppointmentNoteTitleMaxLength = (int)numAppointmentTitleMaxLength.Value;
            _settings.AppointmentDefaultNoteTags = txtAppointmentDefaultTags.Text
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();
            _settings.AppointmentSaveAttachments = chkAppointmentSaveAttachments.Checked;
            _settings.CreateMeetingNotes = chkCreateMeetingNotes.Checked;
            _settings.MeetingNoteTemplate = txtMeetingNoteTemplate.Text;
            _settings.GroupRecurringMeetings = chkGroupRecurringMeetings.Checked;
            _settings.SaveCancelledAppointments = chkSaveCancelledAppointments.Checked;
            _settings.AppointmentTaskCreation = cmbAppointmentTaskCreation.SelectedItem?.ToString() ?? "None";
            _settings.AppointmentTemplateFile = txtAppointmentTemplateFile.Text.Trim();

            // Contacts tab
            _settings.ContactsFolder = txtContactsFolder.Text;
            _settings.EnableContactSaving = chkEnableContactSaving.Checked;
            _settings.SearchEntireVaultForContacts = chkSearchEntireVaultForContacts.Checked;
            _settings.ContactFilenameFormat = txtContactFilenameFormat.Text.Trim();
            _settings.ContactTemplateFile = txtContactTemplateFile.Text.Trim();
            _settings.ContactNoteIncludeDetails = chkContactNoteIncludeDetails.Checked;
            _settings.ContactLinkFormat = txtContactLinkFormat.Text.Trim();
            _settings.ContactDateFormat = txtContactDateFormat.Text.Trim();
            _settings.EmailDateFormat = txtEmailDateFormat.Text.Trim();
            _settings.AppointmentDateFormat = txtAppointmentDateFormat.Text.Trim();

            // Tasks tab
            _settings.CreateObsidianTask = chkCreateObsidianTask.Checked;
            _settings.CreateOutlookTask = chkCreateOutlookTask.Checked;
            _settings.AskForDates = chkAskForDates.Checked;
            _settings.DefaultDueDays = (int)numDefaultDueDays.Value;
            _settings.DefaultReminderDays = (int)numDefaultReminderDays.Value;
            _settings.DefaultReminderHour = (int)numDefaultReminderHour.Value;
            _settings.DefaultTaskTags = txtDefaultTaskTags.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim()).Where(t => !string.IsNullOrEmpty(t)).ToList();
            _settings.TaskTemplateFile = txtTaskTemplateFile.Text.Trim();

            // Threading tab
            _settings.GroupEmailThreads = chkGroupEmailThreads.Checked;
            _settings.MoveDateToFrontInThread = chkMoveDateToFrontInThread.Checked;
            _settings.ThreadTemplateFile = txtThreadTemplateFile.Text.Trim();

            // Attachments tab
            _settings.AttachmentsFolder = txtAttachmentsFolder.Text.Trim();
            _settings.AttachmentStorageMode = (AttachmentStorageMode)cmbAttachmentStorageMode.SelectedIndex;
            _settings.SaveInlineImages = chkSaveInlineImages.Checked;
            _settings.SaveAllAttachments = chkSaveAllAttachments.Checked;
            _settings.UseObsidianWikilinks = chkUseObsidianWikilinks.Checked;

            // Developer tab
            _settings.ShowDevelopmentSettings = chkShowDevelopmentSettings.Checked;
            _settings.ShowThreadDebug = chkShowThreadDebug.Checked;

            // Auto-Sling tab
            _settings.EnableAutoSling = chkEnableAutoSling.Checked;
            _settings.AutoSlingNotificationMode = cmbNotificationMode.SelectedItem?.ToString() ?? "Toast";
            _settings.EnableFlagToSling = chkEnableFlagToSling.Checked;
            _settings.SentToObsidianCategory = txtSentToObsidianCategory.Text.Trim();

            _settings.AutoSlingRules.Clear();
            foreach (DataGridViewRow row in dgvAutoSlingRules.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string ruleType = row.Cells["Type"].Value?.ToString() ?? "Sender";
                string rulePattern = row.Cells["Pattern"].Value?.ToString() ?? string.Empty;
                bool ruleEnabled = row.Cells["Enabled"].Value is bool b && b;
                _settings.AutoSlingRules.Add(new AutoSlingRule { Type = ruleType, Pattern = rulePattern, Enabled = ruleEnabled });
            }

            _settings.WatchedFolders.Clear();
            foreach (DataGridViewRow row in dgvWatchedFolders.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string folderPath = row.Cells["FolderPath"].Value?.ToString() ?? string.Empty;
                string customTemplate = row.Cells["CustomTemplate"].Value?.ToString() ?? string.Empty;
                bool folderEnabled = row.Cells["Enabled"].Value is bool fb && fb;
                _settings.WatchedFolders.Add(new WatchedFolder { FolderPath = folderPath, CustomTemplate = customTemplate, Enabled = folderEnabled });
            }

            _settings.Save();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (InputDialog form = new InputDialog("Add Pattern", "Enter regex pattern:"))
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
                using (InputDialog form = new InputDialog("Edit Pattern", "Edit regex pattern:", lstPatterns.SelectedItem.ToString()))
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
            grpDevelopment.Visible = true;
            chkShowThreadDebug.Visible = chkShowDevelopmentSettings.Checked;
        }

        private void lnkBuyMeACoffee_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkBuyMeACoffee.LinkVisited = true;
            SupportService.OpenBuyMeACoffeeLink(this);
        }
    }
}
