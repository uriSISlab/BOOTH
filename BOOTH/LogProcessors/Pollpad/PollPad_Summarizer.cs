using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.PollPad
{
    class PollPad_Summarizer : LogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "PollPad";

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
