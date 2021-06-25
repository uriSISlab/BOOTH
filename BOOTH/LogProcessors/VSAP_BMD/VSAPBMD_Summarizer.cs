namespace BOOTH.LogProcessors.VSAP_BMD
{
    class VSAPBMD_Summarizer : LogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "VSAPBMD";

        protected override ColumnInfo GetDurationColumnInfo()
        {
            return new ColumnInfo("A", "Duration (mm:ss)");
        }

        protected override ColumnInfo GetEventTypeColumnInfo()
        {
            return new ColumnInfo("C", "Scan Type");
        }
    }
}
