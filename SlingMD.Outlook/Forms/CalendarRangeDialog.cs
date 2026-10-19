using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// Date-range and calendar picker for bulk calendar export. Processing triggered by this dialog
    /// uses <see cref="SlingMD.Outlook.Services.ContactInteractionMode.Automated"/> — no
    /// interactive dialogs are shown during the export run.
    /// </summary>
    public class CalendarRangeDialog : Form
    {
        internal DateTimePicker dtpStart;
        internal DateTimePicker dtpEnd;
        internal CheckedListBox clbCalendars;
        private Label lblCalendars;
        private Button btnOk;
        private Button btnCancel;

        public DateTime StartDate => dtpStart.Value.Date;
        public DateTime EndDate => dtpEnd.Value.Date;
        public bool IsOkEnabled => btnOk.Enabled;

        /// <summary>
        /// The checked calendars, in list order. Empty when no calendars were offered, in which case
        /// the caller falls back to its saved or default calendars.
        /// </summary>
        public List<CalendarFolderInfo> SelectedCalendars => clbCalendars.CheckedItems.Cast<CalendarFolderInfo>().ToList();

        public CalendarRangeDialog()
            : this(null, null)
        {
        }

        /// <param name="calendars">Calendars to offer. Null or empty hides the choice.</param>
        /// <param name="preselected">Calendars to check initially, matched by EntryID.</param>
        public CalendarRangeDialog(IEnumerable<CalendarFolderInfo> calendars, IEnumerable<CalendarFolderInfo> preselected)
        {
            InitializeControls();
            PopulateCalendars(calendars, preselected);
        }

        private void InitializeControls()
        {
            Label lblStart = new Label
            {
                Text = "Start Date:",
                AutoSize = true,
                Location = new Point(12, 15)
            };

            dtpStart = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today,
                Location = new Point(90, 12),
                Width = 150
            };

            Label lblEnd = new Label
            {
                Text = "End Date:",
                AutoSize = true,
                Location = new Point(12, lblStart.Bottom + 20)
            };

            dtpEnd = new DateTimePicker
            {
                Format = DateTimePickerFormat.Short,
                Value = DateTime.Today,
                Location = new Point(90, lblStart.Bottom + 17),
                Width = 150
            };

            lblCalendars = new Label
            {
                Text = "Calendars:",
                AutoSize = true,
                Location = new Point(12, dtpEnd.Bottom + 16)
            };

            clbCalendars = new CheckedListBox
            {
                CheckOnClick = true,
                IntegralHeight = false,
                HorizontalScrollbar = true,
                Location = new Point(12, lblCalendars.Bottom + 4),
                Size = new Size(336, 140)
            };

            btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(clbCalendars.Right - 160, clbCalendars.Bottom + 16),
                Width = 75,
                Enabled = true
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new Point(btnOk.Right + 10, btnOk.Top),
                Width = 75
            };

            dtpStart.ValueChanged += OnDateValueChanged;
            dtpEnd.ValueChanged += OnDateValueChanged;
            clbCalendars.ItemCheck += OnCalendarItemCheck;

            Controls.AddRange(new Control[] { lblStart, dtpStart, lblEnd, dtpEnd, lblCalendars, clbCalendars, btnOk, btnCancel });

            Text = "Save Date Range";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            AcceptButton = btnOk;
            CancelButton = btnCancel;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(clbCalendars.Right + 12, btnOk.Bottom + 16);
        }

        private void PopulateCalendars(IEnumerable<CalendarFolderInfo> calendars, IEnumerable<CalendarFolderInfo> preselected)
        {
            List<CalendarFolderInfo> preselectedList = (preselected ?? Enumerable.Empty<CalendarFolderInfo>())
                .Where(c => c != null)
                .ToList();

            foreach (CalendarFolderInfo calendar in calendars ?? Enumerable.Empty<CalendarFolderInfo>())
            {
                if (calendar == null)
                {
                    continue;
                }

                clbCalendars.Items.Add(calendar, preselectedList.Any(p => p.IsSameFolder(calendar)));
            }

            if (clbCalendars.Items.Count == 0)
            {
                lblCalendars.Text = "Calendars: none found. Your default calendar will be exported.";
                clbCalendars.Enabled = false;
            }

            UpdateOkEnabled(clbCalendars.CheckedItems.Count);
        }

        private void OnDateValueChanged(object sender, EventArgs e)
        {
            UpdateOkEnabled(clbCalendars.CheckedItems.Count);
        }

        private void OnCalendarItemCheck(object sender, ItemCheckEventArgs e)
        {
            // ItemCheck fires before the item's state changes, so CheckedItems still holds the old set.
            int checkedCount = clbCalendars.CheckedItems.Count;
            if (e.CurrentValue != CheckState.Checked && e.NewValue == CheckState.Checked)
            {
                checkedCount++;
            }
            else if (e.CurrentValue == CheckState.Checked && e.NewValue != CheckState.Checked)
            {
                checkedCount--;
            }

            UpdateOkEnabled(checkedCount);
        }

        private void UpdateOkEnabled(int checkedCount)
        {
            bool datesValid = dtpStart.Value.Date <= dtpEnd.Value.Date;
            bool calendarsValid = clbCalendars.Items.Count == 0 || checkedCount > 0;
            btnOk.Enabled = datesValid && calendarsValid;
        }
    }
}
