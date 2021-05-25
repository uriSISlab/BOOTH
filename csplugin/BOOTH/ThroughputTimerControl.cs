using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    public class ThroughputTimerControl : VotingBoothTimerControl
    {
        public ThroughputTimerControl(SheetWriter writer, int number) : base(writer, number)
        {
            base.GetPicture().Image = global::BOOTH.Properties.Resources.Vote_scaled;
            base.GetHeadingLabel().Text = "Throughput " + number; 
            this.writer.Return();
            this.writer.WriteLineArrWithoutLineBreak(new string[] {
                "Throughput " + number + " start",
                "Throughput " + number + " end",
                "Throughput " + number + " duration",
                "Throughput " + number + " comments"
            });
            this.writer.FormatPretty();
        }
    }
}
