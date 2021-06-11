namespace BOOTH
{
    public class BallotScanningTimerControl : VotingBoothTimerControl
    {
        public BallotScanningTimerControl(DynamicSheetWriter writer, int number) : base(writer, number)
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

        protected override string[][] GetHelpTextItems()
        {
            return new string[][] {
                new string[] { Properties.Resources.scannerStartName, Properties.Resources.scannerStartDescription },
                new string[] { Properties.Resources.scannerStopName, Properties.Resources.scannerStopDescription },
                new string[] { Properties.Resources.undoName, Properties.Resources.undoDescription },
                new string[] { Properties.Resources.clearName, Properties.Resources.clearDescription },
            };
        }
    }
}
