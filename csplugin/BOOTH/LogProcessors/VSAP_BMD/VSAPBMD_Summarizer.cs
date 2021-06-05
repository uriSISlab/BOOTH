using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.VSAP_BMD
{
    class VSAPBMD_Summarizer : ILogSummarizer
    {
        public static readonly string MACHINE_TYPE_TAG = "VSAPBMD";

        public void CreateSummaryFrom(Worksheet sheet)
        {
            // PivotCache cache = ThisAddIn.app.ActiveWorkbook.PivotCaches()
            //    .Create(XlPivotTableSourceType.xlDatabase, sheet.UsedRange);
            string outSheetName = "VSAP BMD Summary Statistics";
            Worksheet outSheet = Util.TryAddingSheetWithName(outSheetName);
            for (int i = 1; outSheet == null; i++)
            {
                outSheet = Util.TryAddingSheetWithName(outSheetName + " " + i);
            }
            // outSheet.PivotTables().Add(cache, outSheet.Range["A1"]);
            PivotTable table = outSheet.PivotTableWizard(
                XlPivotTableSourceType.xlDatabase,
                sheet.UsedRange,
                outSheet.Range["A1"],
                "VSAP BMD Log Summary Statistics"
                );
        }
    }
}
