using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class TimerBaseForm : Form
    {
        private readonly List<TimerControl> timers;
        private Worksheet sheet;

        public TimerBaseForm()
        {
            InitializeComponent();
            this.timers = new List<TimerControl>();
            this.storeCommentButton.Enabled = false;
        }

        public static TimerBaseForm CreateForType(Timers.TimerFormType timerFormType)
        {
            TimerBaseForm timerBase = new TimerBaseForm();
            switch (timerFormType)
            {
                case Timers.TimerFormType.CHECKIN:
                    timerBase.SetupCheckinTimersForm(ThisAddIn.app.ActiveWorkbook.ActiveSheet);
                    break;
                case Timers.TimerFormType.CHECKIN_ARRIVAL:
                    timerBase.SetupCheckinArrivalTimersForm(ThisAddIn.app.ActiveWorkbook.ActiveSheet);
                    break;
                case Timers.TimerFormType.VOTING_BOOTH:
                    timerBase.SetupVotingBoothTimersForm(ThisAddIn.app.ActiveWorkbook.ActiveSheet);
                    break;
                case Timers.TimerFormType.BMD:
                    timerBase.SetupBMDTimersForm(ThisAddIn.app.ActiveWorkbook.ActiveSheet);
                    break;
                case Timers.TimerFormType.BALLOT_SCANNING:
                    timerBase.SetupBallotScanningTimersForm(ThisAddIn.app.ActiveWorkbook.ActiveSheet);
                    break;
                case Timers.TimerFormType.THROUGHPUT_ARRIVAL:
                    timerBase.SetupThroughputArrivalTimersForm(ThisAddIn.app.ActiveWorkbook.ActiveSheet);
                    break;
            }
            return timerBase;
        }

        private void RegisterTimer(TimerControl timerControl)
        {
            this.timers.Add(timerControl);
            this.commentTimerSelectCombo.Items.Add(timerControl.GetHeadingText());
        }

        private void SetupCheckinArrivalTimersForm(Worksheet sheet)
        {
            SheetWriter writer = new SheetWriter(sheet, 0, 0);
            ArrivalTimerControl arrivalTimer = new ArrivalTimerControl(writer);
            this.leftPanel.Controls.Add(arrivalTimer);
            this.RegisterTimer(arrivalTimer);
            SetupCheckinTimersForm(sheet, 0, 3);
        }

        private void SetupCheckinTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            Timers.TimerType timerType = Timers.TimerType.CHECKIN;
            Timers.TimerType[] timerTypes = new Timers.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupVotingBoothTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            Timers.TimerType timerType = Timers.TimerType.VOTING_BOOTH;
            Timers.TimerType[] timerTypes = new Timers.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupBMDTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            Timers.TimerType timerType = Timers.TimerType.BMD;
            Timers.TimerType[] timerTypes = new Timers.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupBallotScanningTimersForm(Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            Timers.TimerType timerType = Timers.TimerType.BALLOT_SCANNING;
            Timers.TimerType[] timerTypes = new Timers.TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, rowOffset, columnOffset);
        }

        private void SetupThroughputArrivalTimersForm(Worksheet sheet)
        {
            SheetWriter writer = new SheetWriter(sheet, 0, 0);
            ArrivalTimerControl arrivalTimer = new ArrivalTimerControl(writer);
            this.leftPanel.Controls.Add(arrivalTimer);
            this.RegisterTimer(arrivalTimer);
            Timers.TimerType timerType = Timers.TimerType.THROUGHPUT;
            Timers.TimerType[] timerTypes = new Timers.TimerType[] { timerType, timerType, timerType, timerType, timerType };
            PopulateTimersTablePanel(timerTypes, sheet, 0, 3);
        }

        private void PopulateTimersTablePanel(Timers.TimerType[] timerTypes, Worksheet sheet, int rowOffset = 0, int columnOffset = 0)
        {
            this.timersPanel.ColumnCount = timerTypes.Length;
            this.timersPanel.RowCount = 1;
            this.timersPanel.ColumnStyles.Clear();
            for (int i = 0; i < timerTypes.Length; i++)
            {
                int columnCount = Timers.GetColumnCountForTimerType(timerTypes[i]);
                SheetWriter writer = new SheetWriter(sheet, rowOffset + 0, columnOffset + i * columnCount);
                TimerControl control = Timers.GetTimerControl(timerTypes[i], writer, i + 1);
                this.timersPanel.Controls.Add(control, i, 0);
                this.timersPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f / timerTypes.Length));
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
