namespace BOOTH.LogProcessors.Dominion_ICE
{
    class DICE_Summarizer : LogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "DICE";

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
            return "Event";
        }
    }
}
