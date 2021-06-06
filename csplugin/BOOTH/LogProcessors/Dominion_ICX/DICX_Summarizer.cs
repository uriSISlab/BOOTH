using BOOTH.LogProcessors.Dominion_ICE;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.Dominion_ICX
{
    class DICX_Summarizer : ILogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "DICX";

        public void CreateSummaryFrom(Worksheet sheet)
        {
            (new DICE_Summarizer()).CreateSummaryFrom(sheet);
        }
    }
}
