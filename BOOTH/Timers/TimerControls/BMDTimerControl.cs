using System;
using System.Drawing;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class BMDTimerControl : TimerControl
    {
        private bool neverStarted;
        private DateTime startStamp;
        private bool helped;
        private Color previousColor;

        public BMDTimerControl(DynamicSheetWriter writer, int number) : base(writer, number)
        {
            InitializeComponent();
            this.headingLabel.Text = "BMD " + number;
            this.headingLabel.AutoSize = false;
            this.headingLabel.TextAlign = ContentAlignment.TopCenter;
            this.headingLabel.Dock = DockStyle.Fill;
            this.writer.WriteLineArrWithoutLineBreak(new string[] {
                "BMD " + number + " start",
                "BMD " + number + " end",
                "BMD " + number + " duration",
                "BMD " + number + " help",
                "BMD " + number + " comments"
                });
            this.writer.FormatPretty();
            this.clearButton.Click += (s, e) => { this.textbox.Text = ""; };
            this.stopButton.Enabled = false;
            this.helpButton.Enabled = false;
            this.neverStarted = true;
            this.helped = false;

            // Set keyboard shortcuts
            string[] shortcuts = TimerControl.GetShortCutsForStartAndStop(number);
            if (shortcuts[0] != null && shortcuts[1] != null)
            {
                this.startButton.Text = this.startButton.Text + " (&" + shortcuts[0] + ")";
                this.stopButton.Text = this.stopButton.Text + " (&" + shortcuts[1] + ")";
            }
        }

        public override string GetHeadingText()
        {
            return this.headingLabel.Text;
        }

        public override void AddComment(string comment)
        {
            if (this.neverStarted)
            {
                this.writer.LineBreak();
            }
            this.writer.WriteLineArrWithoutLineBreak(new string[] { null, null, null, null, comment });
            this.writer.Return();
            if (this.neverStarted)
            {
                this.writer.PreviousLine();
            }
        }

        private void Reset()
        {
            this.startButton.Enabled = true;
            this.stopButton.Enabled = false;
            this.helped = false;
            this.helpButton.Enabled = false;
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            this.writer.LineBreak();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { now.ToString() }, new FieldType[] { FieldType.DATETIME });
            this.writer.Return();
            this.startButton.Enabled = false;
            this.stopButton.Enabled = true;
            this.startStamp = now;
            this.neverStarted = false;
            this.undoLastButton.Enabled = true;
            this.helpButton.Enabled = true;
            this.previousColor = this.BackColor;
            this.BackColor = Color.LightGreen;
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            string duration = (now - this.startStamp).ToString(@"mm\:ss");
            this.writer.Return();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { null, now.ToString(), duration, this.helped ? "Helped" : "" },
                new FieldType[] { FieldType.STRING, FieldType.DATETIME, FieldType.TIMESPAN_MMSS, FieldType.STRING });
            this.writer.Return();
            this.BackColor = this.previousColor;
            Reset();
        }

        private void UndoLastButton_Click(object sender, EventArgs e)
        {
            if (this.neverStarted)
            {
                return;
            }
            this.writer.Return();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { "", "", "", "", "" });
            this.writer.PreviousLine();
            this.Reset();
            if (this.writer.GetRowNum() == 1)
            {
                this.neverStarted = true;
                this.undoLastButton.Enabled = false;
            }
        }

        private void HelpButton_Click(object sender, EventArgs e)
        {
            this.helped = true;
            this.helpButton.Enabled = false;
        }

        protected override string[][] GetHelpTextItems()
        {
            return new string[][] {
                new string[] { Properties.Resources.bmdStartName, Properties.Resources.bmdStartDescription },
                new string[] { Properties.Resources.bmdStopName, Properties.Resources.bmdStopDescription },
                new string[] { Properties.Resources.bmdHelpName, Properties.Resources.bmdHelpDescription },
                new string[] { Properties.Resources.undoName, Properties.Resources.undoDescription },
                new string[] { Properties.Resources.clearName, Properties.Resources.clearDescription },
            };
        }

        private void ShowHelpButton_Click(object sender, EventArgs e)
        {
            this.OpenHelpForm();
        }
    }
}
