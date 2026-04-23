using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public class CalendarRangeDialog : Form
    {
        internal DateTimePicker dtpStart;
        internal DateTimePicker dtpEnd;
        private Button btnOk;
        private Button btnCancel;

        public DateTime StartDate => dtpStart.Value.Date;
        public DateTime EndDate => dtpEnd.Value.Date;
        public bool IsOkEnabled => btnOk.Enabled;

        public CalendarRangeDialog()
        {
            InitializeControls();
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

            btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(90, dtpEnd.Bottom + 16),
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

            Controls.AddRange(new Control[] { lblStart, dtpStart, lblEnd, dtpEnd, btnOk, btnCancel });

            Text = "Select Date Range";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            AcceptButton = btnOk;
            CancelButton = btnCancel;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(260, btnOk.Bottom + 16);
        }

        private void OnDateValueChanged(object sender, EventArgs e)
        {
            btnOk.Enabled = dtpStart.Value.Date <= dtpEnd.Value.Date;
        }
    }
}
