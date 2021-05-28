using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    public partial class ArrivalTimerControl : TimerControl
    {

        private bool neverStarted;
        private int totalArrivals;

        public ArrivalTimerControl(SheetWriter writer) : base(writer, 0)
        {
            InitializeComponent();
            writer.WriteLineArrWithoutLineBreak(new string[] {
                "Arrival time",
                "Arrival type",
                "Arrival comment"
                });
            writer.FormatPretty();
            this.neverStarted = true;
            this.undoLastButton.Enabled = false;
            this.totalArrivals = 0;
        }

        public override string GetHeadingText()
        {
            return "Arrival Timer";
        }

        public override void AddComment(string comment)
        {
            if (this.neverStarted)
            {
                MessageBox.Show("Please record at least one arrival before adding comment to it.");
                return;
            }
            this.writer.Return();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { null, null, comment });
            this.writer.Return();
        }

        private void IncrementArrivalCount()
        {
            this.totalArrivals++;
            this.arrivalCountLabel.Text = totalArrivals.ToString();
        }
        private void DecrementArrivalCount()
        {
            this.totalArrivals--;
            this.arrivalCountLabel.Text = totalArrivals.ToString();
        }


        private void ArrivalButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            this.writer.LineBreak();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { now.ToString(), "Normal" }, new FieldType[] { FieldType.DATETIME,
                FieldType.STRING });
            this.neverStarted = false;
            this.undoLastButton.Enabled = true;
            this.IncrementArrivalCount();
        }

        private void VbmArrivalButton_Click(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            this.writer.LineBreak();
            this.writer.WriteLineArrWithoutLineBreak(new string[] { now.ToString(), "VBM" }, new FieldType[] { FieldType.DATETIME,
                FieldType.STRING });
            this.neverStarted = false;
            this.undoLastButton.Enabled = true;
            this.IncrementArrivalCount();
        }

        private void UndoLastButton_Click(object sender, EventArgs e)
        {
            if (!this.neverStarted)
            {
                this.writer.Return();
                this.writer.WriteLineArrWithoutLineBreak(new string[] { "", "", "", "", "" });
                this.writer.PreviousLine();
                this.DecrementArrivalCount();
                if (this.writer.GetRowNum() == 1)
                {
                    this.neverStarted = true;
                    this.undoLastButton.Enabled = false;
                }
            }
        }

        private void HelpButton_Click(object sender, EventArgs e)
        {
            this.OpenHelpForm();
        }

        protected override string[][] GetHelpTextItems()
        {
            return new string[][] {
                new string[] { Properties.Resources.arrivalName, Properties.Resources.arrivalDescription },
                new string[] { Properties.Resources.vbmArrivalName, Properties.Resources.vbmArrivalDescription },
            };
        }
    }
}
