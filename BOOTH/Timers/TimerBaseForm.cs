using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class TimerBaseForm : Form
    {
        public enum TimerFormType
        {
            CHECKIN,
            CHECKIN_ARRIVAL,
            VOTING_BOOTH,
            BMD,
            BALLOT_SCANNING,
            THROUGHPUT_ARRIVAL
        }

        private readonly List<TimerControl> timers;

        public TimerBaseForm()
        {
            InitializeComponent();
            this.timers = new List<TimerControl>();
            this.storeCommentButton.Enabled = false;
        }

        public static TimerBaseForm CreateForType(TimerFormType timerFormType, Worksheet sheet, string title)
        {
            TimerBaseForm timerBase = new TimerBaseForm { Text = title };
            switch (timerFormType)
            {
                case TimerFormType.CHECKIN:
                    timerBase.SetupCheckinTimersForm(sheet);
                    break;
                case TimerFormType.CHECKIN_ARRIVAL:
                    timerBase.SetupCheckinArrivalTimersForm(sheet);
                    break;
                case TimerFormType.VOTING_BOOTH:
                    timerBase.SetupVotingBoothTimersForm(sheet);
                    break;
                case TimerFormType.BMD:
                    timerBase.SetupBMDTimersForm(sheet);
                    break;
                case TimerFormType.BALLOT_SCANNING:
                    timerBase.SetupBallotScanningTimersForm(sheet);
                    break;
                case TimerFormType.THROUGHPUT_ARRIVAL:
                    timerBase.SetupThroughputArrivalTimersForm(sheet);
                    break;
            }
            return timerBase;
        }

        public static TimerBaseForm CreateWithTimerTypes(TimerControl.TimerType[] timerTypes,
            bool includeArrival, Worksheet sheet, string title)
        {
            TimerBaseForm baseForm = new TimerBaseForm { Text = title };
            if (includeArrival)
            {
                baseForm.PopulateArrivalTimer(sheet);
            }
            baseForm.PopulateTimersTablePanel(timerTypes, sheet, 0, includeArrival ? 3 : 0);
            return baseForm;
        }

        private void RegisterTimer(TimerControl timerControl)
        {
            this.timers.Add(timerControl);
            this.commentTimerSelectCombo.Items.Add(timerControl.GetHeadingText());
        }

        private void SetupCheckinArrivalTimersForm(Worksheet sheet)
        {
            PopulateArrivalTimer(sheet);
            SetupCheckinTimersForm(sheet, 0, 3);
        }

        private void SetupCheckinTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            TimerControl.TimerType timerType = TimerControl.TimerType.CHECKIN;
            TimerControl.TimerType[] timerTypes = new TimerControl.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupVotingBoothTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            TimerControl.TimerType timerType = TimerControl.TimerType.VOTING_BOOTH;
            TimerControl.TimerType[] timerTypes = new TimerControl.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupBMDTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            TimerControl.TimerType timerType = TimerControl.TimerType.BMD;
            TimerControl.TimerType[] timerTypes = new TimerControl.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupBallotScanningTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            TimerControl.TimerType timerType = TimerControl.TimerType.BALLOT_SCANNING;
            TimerControl.TimerType[] timerTypes = new TimerControl.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupThroughputArrivalTimersForm(Worksheet sheet)
        {
            PopulateArrivalTimer(sheet);
            TimerControl.TimerType timerType = TimerControl.TimerType.THROUGHPUT;
            TimerControl.TimerType[] timerTypes = new TimerControl.TimerType[] { timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, 0, 3);
        }

        private void PopulateArrivalTimer(Worksheet sheet)
        {
            DynamicSheetWriter writer = new DynamicSheetWriter(sheet, 0, 0);
            ArrivalTimerControl arrivalTimer = new ArrivalTimerControl(writer);
            this.leftPanel.Controls.Add(arrivalTimer);
            this.RegisterTimer(arrivalTimer);
        }

        private void PopulateTimersTablePanel(TimerControl.TimerType[] timerTypes, Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            this.timersPanel.ColumnCount = timerTypes.Length;
            this.timersPanel.RowCount = 1;
            this.timersPanel.ColumnStyles.Clear();
            for (int i = 0; i < timerTypes.Length; i++)
            {
                int columnCount = i > 0 ? TimerControl.GetColumnCountForTimerType(timerTypes[i - 1]) : 0;
                DynamicSheetWriter writer = new DynamicSheetWriter(sheet, rowOffset + 0, columnOffset + i * columnCount);
                TimerControl control = TimerControl.GetTimerControl(timerTypes[i], writer, i + 1);
                this.timersPanel.Controls.Add(control, i, 0);
                if (timerTypes.Length <= 6)
                {
                    this.timersPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / timerTypes.Length));
                }
                this.RegisterTimer(control);
            }
        }

        private void StoreCommentButton_Click(object sender, EventArgs e)
        {
            TimerControl timer = this.timers[this.commentTimerSelectCombo.SelectedIndex];
            timer.AddComment(this.commentTextBox.Text);
            this.commentTextBox.Clear();
        }

        private void CommentTimerSelectCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.storeCommentButton.Enabled = this.commentTimerSelectCombo.SelectedIndex >= 0;
        }

        private void SaveWorksheetButton_Click(object sender, EventArgs e)
        {
            ThisAddIn.app.ActiveWorkbook.Save();
        }
    }
}
