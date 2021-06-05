using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors
{
    public interface ILogSummarizer
    {
        void CreateSummaryFrom(Worksheet sheet);
    }
}
