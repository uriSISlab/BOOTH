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
    public class BallotScanningTimerControl : VotingBoothTimerControl
    {
        public BallotScanningTimerControl(SheetWriter writer, int number) : base(writer, number)
        {
            base.GetPicture().Image = global::BOOTH.Properties.Resources.DS200_BallotBox_resized;
            base.GetHeadingLabel().Text = "Ballot Scanning " + number; 
            this.writer.Return();
            this.writer.WriteLineArrWithoutLineBreak(new string[] {
                "Ballot Scanner " + number + " start",
                "Ballot Scanner " + number + " end",
                "Ballot Scanner " + number + " duration",
                "Ballot Scanner " + number + " comments"
            });
            this.writer.FormatPretty();
        }
    }
}
