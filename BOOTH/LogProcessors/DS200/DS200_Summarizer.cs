namespace BOOTH.LogProcessors.DS200
{
    class DS200_Summarizer : LogSummarizer
    {

        public static readonly string MACHINE_TYPE_TAG = "DS200";

        protected override ColumnInfo GetDurationColumnInfo()
        {
            return new ColumnInfo("A", "Duration (mm:ss)");
        }

        protected override ColumnInfo GetEventTypeColumnInfo()
        {
            return new ColumnInfo("C", "Scan Type");
        }

        protected override ColumnInfo[] GetCategoricalColumnInfos()
        {
            return new ColumnInfo[]
            {
                new ColumnInfo("D", "Ballot Cast Status"),
                new ColumnInfo("F", "Ballot Type"),
            };
        }
    }
}
