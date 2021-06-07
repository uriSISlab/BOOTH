using System;
using System.Drawing;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class VotingBoothTimerControl : TimerControl
    {
        private bool neverStarted;
        private DateTime startStamp;
        private Color previousColor;

        public VotingBoothTimerControl() : base(null, 0)
        {
            // NOTE: This constructor is here because the UI designer needs the base class
            // of a UI element to be non-abstract and to have a no-argument constructor.
            // (BallotScanningTimerControl and ThroughputTimerControl inherit from it)
            // This constructor should not be used in practice.
        }

        public VotingBoothTimerControl(SheetWriter writer, int number) : base(writer, number)
        {
            InitializeComponent();
            this.headingLabel.Text = "Vooting Booth " + number;
            this.headingLabel.AutoSize = false;
            this.headingLabel.TextAlign = ContentAlignment.TopCenter;
            this.headingLabel.Dock = DockStyle.Fill;
            this.writer.WriteLineArrWithoutLineBreak(new string[] {
                "Voting Booth " + number + " start",
                "Voting Booth " + number + " end",
                "Voting Booth " + number + " duration",
                "Voting Booth " + number + " comments"
                });
            this.writer.FormatPretty();
            this.clearButton.Click += (s, e) => { this.textbox.Text = ""; };
            this.startButton.Click += StartButton_Click;
            this.stopButton.Click += StopButton_Click;
            this.stopButton.Enabled = false;
            this.neverStarted = true;

            // Set keyboard shortcuts
            string[] shortcuts = TimerControl.GetShortCutsForStartAndStop(number);
            if (shortcuts[0] != null && shortcuts[1] != null)
            {
                this.startButton.Text = this.startButton.Text + " (&" + shortcuts[0] + ")";
                this.stopButton.Text = this.stopButton.Text + " (&" + shortcuts[1] + ")";
            }
        }

        protected PictureBox GetPicture()
        {
            return this.pictureBox1;
        }

        protected Label GetHeadingLabel()
        {
            return this.headingLabel;
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
            this.writer.WriteLineArrWithoutLineBreak(new string[] { null, null, null, comment });
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
            this.previousColor = this.BackColor;
            this.BackColor = Color.LightGreen;
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            string duration = (now - this.startStamp).ToString(@"mm\:ss");
            this.writer.Return();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { null, now.ToString(), duration },
                new FieldType[] { FieldType.STRING, FieldType.DATETIME, FieldType.TIMESPAN_MMSS });
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
            this.writer.WriteLineArrWithoutLineBreak(new string[] { "", "", "", "" });
            this.writer.PreviousLine();
            this.Reset();
            if (this.writer.GetRowNum() == 1)
            {
                this.neverStarted = true;
                this.undoLastButton.Enabled = false;
            }
        }

        private void Helpbutton_Click(object sender, EventArgs e)
        {
            this.OpenHelpForm();
        }

        protected override string[][] GetHelpTextItems()
        {
            return new string[][] {
                new string[] { Properties.Resources.votingBoothStartName, Properties.Resources.votingBoothStartDescription },
                new string[] { Properties.Resources.votingBoothStopName, Properties.Resources.votingBoothStopDescription },
                new string[] { Properties.Resources.undoName, Properties.Resources.undoDescription },
                new string[] { Properties.Resources.clearName, Properties.Resources.clearDescription },
            };
        }
    }
}
