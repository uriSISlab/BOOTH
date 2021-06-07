using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Security.Policy;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace BOOTH
{
    public partial class CheckInTimerControl : TimerControl
    {

        private DateTime startStamp;
        private string checkInType;
        private bool neverStarted;
        private Color previousColor;

        public CheckInTimerControl(SheetWriter writer, int number) : base(writer, number)
        {
            InitializeComponent();
            this.heading.Text = "Check in " + number;
            this.heading.AutoSize = false;
            this.heading.TextAlign = ContentAlignment.TopCenter;
            this.heading.Dock = DockStyle.Fill;
            writer.WriteLineArrWithoutLineBreak(new string[] {
                "Checkin " + number + " start",
                "Checkin " + number + " end",
                "Checkin " + number + " duration",
                "Checkin " + number + " type",
                "Checkin " + number + " comments"
                });
            writer.FormatPretty();
            this.clearButton.Click += (s, e) => { this.textbox.Text = ""; };
            this.startButton.Click += StartButton_Click;
            this.stopButton.Click += StopButton_Click;
            this.stopButton.Enabled = false;
            this.checkInType = "Normal";
            this.neverStarted = true;

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
            return this.heading.Text;
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
            this.vbmButton.Enabled = true;
            this.startButton.Enabled = true;
            this.stopButton.Enabled = false;
            this.startProvButton.Enabled = true;
            this.endProvButton.Enabled = true;
        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            writer.LineBreak();
            writer.WriteLineArrWithoutLineBreak(new string[] { now.ToString() },  new FieldType[] { FieldType.DATETIME });
            writer.Return();
            this.startButton.Enabled = false;
            this.stopButton.Enabled = true;
            this.startStamp = now;
            this.checkInType = "Normal";
            this.neverStarted = false;
            this.previousColor = this.BackColor;
            this.BackColor = Color.LightGreen;
        }

        private void StopButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            string duration = (now - this.startStamp).ToString(@"mm\:ss");
            writer.WriteLineArrWithoutLineBreak(new string[] { null, now.ToString(), duration, this.checkInType },
                new FieldType[] { FieldType.STRING, FieldType.DATETIME, FieldType.TIMESPAN_MMSS });
            writer.Return();
            this.BackColor = this.previousColor;
            Reset();
        }

        private void VbmButton_Click(object sender, EventArgs e)
        {
            this.checkInType = "VBM";
            this.vbmButton.Enabled = false;
            this.startProvButton.Enabled = false;
            this.endProvButton.Enabled = false;
        }

        private void StartProvButton_Click(object sender, EventArgs e)
        {
            this.checkInType = "Given Provisional";
            this.vbmButton.Enabled = false;
            this.startProvButton.Enabled = false;
            this.endProvButton.Enabled = false;
        }

        private void EndProvButton_Click(object sender, EventArgs e)
        {
            this.checkInType = "Returned Provisional";
            this.vbmButton.Enabled = false;
            this.startProvButton.Enabled = false;
            this.endProvButton.Enabled = false;
        }

        private void UndoLastButton_Click(object sender, EventArgs e)
        {
            if (neverStarted)
            {
                return;
            }
            if (this.writer.GetRowNum() == 1)
            {
                return;
            }
            writer.Return();
            writer.WriteLineArrWithoutLineBreak(new string[] {"", "", "", "", ""});
            writer.PreviousLine();
        }

        protected override string[][] GetHelpTextItems()
        {
            return new string[][] {
                new string[] { Properties.Resources.checkinStartName, Properties.Resources.checkinStartDescription },
                new string[] { Properties.Resources.checkinStopName, Properties.Resources.checkinStopDescription },
                new string[] { Properties.Resources.startProvName, Properties.Resources.startProvDescription },
                new string[] { Properties.Resources.endProvName, Properties.Resources.endProvDescription },
                new string[] { Properties.Resources.vbmName, Properties.Resources.vbmDescription },
                new string[] { Properties.Resources.undoName, Properties.Resources.undoDescription },
                new string[] { Properties.Resources.clearName, Properties.Resources.clearDescription },
            };
        }

        private void HelpButton_Click(object sender, EventArgs e)
        {
            this.OpenHelpForm();
        }
    }
}
