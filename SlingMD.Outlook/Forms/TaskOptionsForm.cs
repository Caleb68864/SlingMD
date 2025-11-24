using System;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace SlingMD.Outlook.Forms
{
    public partial class TaskOptionsForm : Form
    {
        private const int formWidth = 550;
        private const int formHeight = 250;
        private const int labelX = 20;
        private const int controlX = 180;
        private const int helpTextX = 350;
        private const int startY = 20;
        private const int lineHeight = 35;

        private Label lblDueDays;
        private NumericUpDown numDueDays;
        private DateTimePicker dtpDueDate;
        private Label lblDueDaysHelp;
        private Label lblReminderDays;
        private NumericUpDown numReminderDays;
        private DateTimePicker dtpReminderDate;
        private Label lblReminderDaysHelp;
        private Label lblReminderHour;
        private NumericUpDown numReminderHour;
        private Label lblReminderHourHelp;
        private CheckBox chkUseRelativeReminder;
        private Button btnOK;
        private Button btnCancel;

        public DateTime DueDate { get; private set; }
        public DateTime ReminderDate { get; private set; }

        public int DueDays => chkUseRelativeReminder.Checked ? (int)numDueDays.Value : (dtpDueDate.Value.Date - DateTime.Now.Date).Days;
        public int ReminderDays => chkUseRelativeReminder.Checked ? (int)numReminderDays.Value : (dtpReminderDate.Value.Date - DateTime.Now.Date).Days;
        public int ReminderHour => (int)numReminderHour.Value;
        public bool UseRelativeReminder => chkUseRelativeReminder.Checked;

        public TaskOptionsForm(int defaultDueDays, int defaultReminderDays, int defaultReminderHour, bool useRelativeReminder = false)
        {
            InitializeComponent();
            DueDate = DateTime.Now.Date.AddDays(defaultDueDays);
            ReminderDate = DateTime.Now.Date.AddDays(defaultReminderDays);
            chkUseRelativeReminder.Checked = useRelativeReminder;
            numDueDays.Value = defaultDueDays;
            numReminderDays.Value = defaultReminderDays;
            numReminderHour.Value = defaultReminderHour;
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        public TaskOptionsForm(DateTime defaultDueDate, DateTime defaultReminderDate)
        {
            InitializeComponent();
            DueDate = defaultDueDate;
            ReminderDate = defaultReminderDate;
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        private void InitializeControls()
        {
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TaskOptionsForm));
            this.SuspendLayout();

            // Initialize all controls
            this.lblDueDays = new Label();
            this.numDueDays = new NumericUpDown();
            this.dtpDueDate = new DateTimePicker();
            this.lblDueDaysHelp = new Label();
            this.lblReminderDays = new Label();
            this.numReminderDays = new NumericUpDown();
            this.dtpReminderDate = new DateTimePicker();
            this.lblReminderDaysHelp = new Label();
            this.lblReminderHour = new Label();
            this.numReminderHour = new NumericUpDown();
            this.lblReminderHourHelp = new Label();
            this.chkUseRelativeReminder = new CheckBox();
            this.btnOK = new Button();
            this.btnCancel = new Button();

            ((System.ComponentModel.ISupportInitialize)(this.numDueDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numReminderDays)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numReminderHour)).BeginInit();

            // lblDueDays
            this.lblDueDays.Location = new Point(labelX, startY);
            this.lblDueDays.Size = new Size(150, 23);
            this.lblDueDays.Text = "Due Date:";
            this.lblDueDays.TextAlign = ContentAlignment.MiddleLeft;

            // numDueDays
            this.numDueDays.Location = new Point(controlX, startY);
            this.numDueDays.Size = new Size(150, 23);
            this.numDueDays.Maximum = 365;
            this.numDueDays.Minimum = 0;

            // dtpDueDate
            this.dtpDueDate.Location = new Point(controlX, startY);
            this.dtpDueDate.Size = new Size(150, 23);
            this.dtpDueDate.Format = DateTimePickerFormat.Short;

            // lblDueDaysHelp
            this.lblDueDaysHelp.Location = new Point(helpTextX, startY);
            this.lblDueDaysHelp.Size = new Size(180, 23);
            this.lblDueDaysHelp.TextAlign = ContentAlignment.MiddleLeft;
            this.lblDueDaysHelp.ForeColor = Color.Gray;

            // lblReminderDays
            this.lblReminderDays.Location = new Point(labelX, startY + lineHeight);
            this.lblReminderDays.Size = new Size(150, 23);
            this.lblReminderDays.Text = "Reminder:";
            this.lblReminderDays.TextAlign = ContentAlignment.MiddleLeft;

            // numReminderDays
            this.numReminderDays.Location = new Point(controlX, startY + lineHeight);
            this.numReminderDays.Size = new Size(150, 23);
            this.numReminderDays.Maximum = 365;
            this.numReminderDays.Minimum = 0;

            // dtpReminderDate
            this.dtpReminderDate.Location = new Point(controlX, startY + lineHeight);
            this.dtpReminderDate.Size = new Size(150, 23);
            this.dtpReminderDate.Format = DateTimePickerFormat.Short;

            // lblReminderDaysHelp
            this.lblReminderDaysHelp.Location = new Point(helpTextX, startY + lineHeight);
            this.lblReminderDaysHelp.Size = new Size(180, 23);
            this.lblReminderDaysHelp.TextAlign = ContentAlignment.MiddleLeft;
            this.lblReminderDaysHelp.ForeColor = Color.Gray;

            // lblReminderHour
            this.lblReminderHour.Location = new Point(labelX, startY + lineHeight * 2);
            this.lblReminderHour.Size = new Size(150, 23);
            this.lblReminderHour.Text = "Reminder Time (Hour):";
            this.lblReminderHour.TextAlign = ContentAlignment.MiddleLeft;

            // numReminderHour
            this.numReminderHour.Location = new Point(controlX, startY + lineHeight * 2);
            this.numReminderHour.Size = new Size(150, 23);
            this.numReminderHour.Maximum = 23;
            this.numReminderHour.Minimum = 0;

            // lblReminderHourHelp
            this.lblReminderHourHelp.Location = new Point(helpTextX, startY + lineHeight * 2);
            this.lblReminderHourHelp.Size = new Size(180, 23);
            this.lblReminderHourHelp.Text = "(24-hour format)";
            this.lblReminderHourHelp.TextAlign = ContentAlignment.MiddleLeft;
            this.lblReminderHourHelp.ForeColor = Color.Gray;

            // chkUseRelativeReminder
            this.chkUseRelativeReminder.Location = new Point(labelX, startY + lineHeight * 3);
            this.chkUseRelativeReminder.Size = new Size(350, 23);
            this.chkUseRelativeReminder.Text = "Use relative dates (days from now)";
            this.chkUseRelativeReminder.CheckedChanged += ChkUseRelativeReminder_CheckedChanged;

            // btnOK
            this.btnOK.Location = new Point(formWidth - 200, formHeight - 50);
            this.btnOK.Size = new Size(75, 23);
            this.btnOK.Text = "OK";
            this.btnOK.DialogResult = DialogResult.OK;
            this.btnOK.Click += BtnOK_Click;

            // btnCancel
            this.btnCancel.Location = new Point(formWidth - 110, formHeight - 50);
            this.btnCancel.Size = new Size(75, 23);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.DialogResult = DialogResult.Cancel;

            // TaskOptionsForm
            this.ClientSize = new Size(formWidth, formHeight);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "TaskOptionsForm";
            this.Text = "Task Options";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.AcceptButton = this.btnOK;
            this.CancelButton = this.btnCancel;

            // Add controls to form
            this.Controls.Add(this.lblDueDays);
            this.Controls.Add(this.numDueDays);
            this.Controls.Add(this.dtpDueDate);
            this.Controls.Add(this.lblDueDaysHelp);
            this.Controls.Add(this.lblReminderDays);
            this.Controls.Add(this.numReminderDays);
            this.Controls.Add(this.dtpReminderDate);
            this.Controls.Add(this.lblReminderDaysHelp);
            this.Controls.Add(this.lblReminderHour);
            this.Controls.Add(this.numReminderHour);
            this.Controls.Add(this.lblReminderHourHelp);
            this.Controls.Add(this.chkUseRelativeReminder);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnCancel);

            ((System.ComponentModel.ISupportInitialize)(this.numDueDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numReminderDays)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numReminderHour)).EndInit();
            this.ResumeLayout(false);
        }

        private void ChkUseRelativeReminder_CheckedChanged(object sender, EventArgs e)
        {
            UpdateControlsVisibility();
            UpdateHelpText();
        }

        private void UpdateControlsVisibility()
        {
            bool useRelative = chkUseRelativeReminder.Checked;
            
            numDueDays.Visible = useRelative;
            dtpDueDate.Visible = !useRelative;
            
            numReminderDays.Visible = useRelative;
            dtpReminderDate.Visible = !useRelative;

            // When switching to absolute dates, update the date pickers based on current numeric values
            if (!useRelative)
            {
                dtpDueDate.Value = DateTime.Now.Date.AddDays((double)numDueDays.Value);
                dtpReminderDate.Value = DateTime.Now.Date.AddDays((double)numReminderDays.Value);
            }
            // When switching to relative dates, update the numeric values based on current date pickers
            else
            {
                numDueDays.Value = Math.Max(0, (dtpDueDate.Value.Date - DateTime.Now.Date).Days);
                numReminderDays.Value = Math.Max(0, (dtpReminderDate.Value.Date - DateTime.Now.Date).Days);
            }
        }

        private void UpdateHelpText()
        {
            if (chkUseRelativeReminder.Checked)
            {
                lblDueDaysHelp.Text = "(Days from today)";
                lblReminderDaysHelp.Text = "(Days before due date)";
            }
            else
            {
                lblDueDaysHelp.Text = "(Select date)";
                lblReminderDaysHelp.Text = "(Select reminder date)";
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (chkUseRelativeReminder.Checked)
            {
                if (numReminderDays.Value > numDueDays.Value)
                {
                    MessageBox.Show(
                        "Reminder days cannot be greater than due days when using relative dates.",
                        "Invalid Reminder",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    this.DialogResult = DialogResult.None;
                }
            }
            else
            {
                if (dtpReminderDate.Value.Date > dtpDueDate.Value.Date)
                {
                    MessageBox.Show(
                        "Reminder date cannot be after the due date.",
                        "Invalid Reminder",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    this.DialogResult = DialogResult.None;
                }
            }
        }
    }
} 