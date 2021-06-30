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

        protected override ColumnInfo GetDurationColumnInfo()
        {
            return new ColumnInfo("A", "Duration (mm:ss)");
        }

        protected override ColumnInfo GetEventTypeColumnInfo()
        {
            return new ColumnInfo("D", "Event");
        }

        protected override ColumnInfo GetTimestampColumnInfo()
        {
            return new ColumnInfo("C", "End Timestamp");
        }

        protected override ColumnInfo[] GetCategoricalColumnInfos()
        {
            return new ColumnInfo[]
            {
                new ColumnInfo("E", "Lookup Method"),
                new ColumnInfo("G", "VBM Cancelled"),
                new ColumnInfo("H", "Assistance Required")
            };
        }
    }
}
