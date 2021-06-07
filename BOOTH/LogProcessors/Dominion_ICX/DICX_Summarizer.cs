using BOOTH.LogProcessors.Dominion_ICE;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.Dominion_ICX
{
    class DICX_Summarizer : LogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "DICX";

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
