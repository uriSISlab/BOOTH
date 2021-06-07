namespace BOOTH.LogProcessors.DS200
{
    class DS200_Summarizer : LogSummarizer
    {

        public static readonly string MACHINE_TYPE_TAG = "DS200";

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
