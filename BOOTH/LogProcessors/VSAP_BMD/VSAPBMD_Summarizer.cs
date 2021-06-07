namespace BOOTH.LogProcessors.VSAP_BMD
{
    class VSAPBMD_Summarizer : LogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "VSAPBMD";

        protected override string GetDurationFieldColumn()
        {
            return "A";
        }

        protected override string GetDurationFieldName()
        {
            return "Duration (mm:ss)";
        }

        protected override string GetEventTypeFieldName()
        {
            return "Scan Type";
        }
    }
}
