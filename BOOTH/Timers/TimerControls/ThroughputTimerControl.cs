namespace BOOTH
{
    public class ThroughputTimerControl : VotingBoothTimerControl
    {
        public ThroughputTimerControl(DynamicSheetWriter writer, int number) : base(writer, number)
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

        protected override string[][] GetHelpTextItems()
        {
            return new string[][] {
                new string[] { Properties.Resources.throughputStartName, Properties.Resources.throughputStartDescription },
                new string[] { Properties.Resources.throughputStopName, Properties.Resources.throughputStopDescription },
                new string[] { Properties.Resources.undoName, Properties.Resources.undoDescription },
                new string[] { Properties.Resources.clearName, Properties.Resources.clearDescription },
            };
        }
    }
}
