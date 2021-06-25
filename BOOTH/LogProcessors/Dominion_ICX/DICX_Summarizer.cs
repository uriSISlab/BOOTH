namespace BOOTH.LogProcessors.Dominion_ICX
{
    class DICX_Summarizer : LogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "DICX";

        protected override ColumnInfo GetDurationColumnInfo()
        {
            return new ColumnInfo("A", "Duration (mm:ss)");
        }

        protected override ColumnInfo GetEventTypeColumnInfo()
        {
            return new ColumnInfo("C", "Event");
        }
    }
}
