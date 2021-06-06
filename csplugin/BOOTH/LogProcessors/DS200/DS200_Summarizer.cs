using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.DS200
{
    class DS200_Summarizer : ILogSummarizer
    {

        public static readonly string MACHINE_TYPE_TAG = "DS200";

        public void CreateSummaryFrom(Worksheet sheet)
        {
            Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            // Create sheet for output
            string suffix = ".. Stats";
            string prefix = Util.Clip(sheet.Name, (31 - suffix.Length));
            string outSheetName = prefix + suffix;
            Worksheet outSheet = Util.TryAddingSheetWithName(outSheetName);
            for (int i = 1; outSheet == null && i < 100; i++)
            {
                string name = Util.Clip(outSheetName, 28) + " " + i;
                outSheet = Util.TryAddingSheetWithName(name);
            }
            if (outSheet == null)
            {
                Util.MessageBox("Could not create new sheet for summary statistics!");
                return;
            }

            sheet.Range["A:A"].NumberFormat = "mm:ss";
            sheet.Range["D:D"].NumberFormat = "general";

            // Create Pivot Cache from Source Data
            PivotCache pvtCache = workbook.PivotCaches().Create(SourceType: XlPivotTableSourceType.xlDatabase,
                //SourceData: srcData);
                SourceData: sheet.UsedRange);

            // Create Pivot table from Pivot Cache
            //PivotTable pvt = pvtCache.CreatePivotTable(TableDestination: startPvt, TableName: "PivotTable2");
            PivotTable pvt = pvtCache.CreatePivotTable(TableDestination: outSheet.Range["A3"], TableName: "PivotTable2");

            pvt.AddDataField(pvt.PivotFields("Scan Type"), "Count of Scan Type", XlConsolidationFunction.xlCount);
            pvt.AddDataField(pvt.PivotFields("Scan Type"), "Percent of Scan Type", XlConsolidationFunction.xlCount);
            pvt.PivotFields("Percent of Scan Type").Calculation = XlPivotFieldCalculation.xlPercentOfColumn;
            pvt.AddDataField(pvt.PivotFields("Duration (mm:ss)"), "Average Duration of Scan Type", XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields("Duration (mm:ss)"), "Max Duration of Scan Type", XlConsolidationFunction.xlMax);
            pvt.AddDataField(pvt.PivotFields("Duration (mm:ss)"), "Standard Deviation of Scan Type", XlConsolidationFunction.xlStDev);
            pvt.PivotFields("Average Duration of Scan Type").NumberFormat = "mm:ss";
            pvt.PivotFields("Max Duration of Scan Type").NumberFormat = "mm:ss";
            pvt.PivotFields("Standard Deviation of Scan Type").NumberFormat = "mm:ss";


            pvt.PivotFields("Scan Type").Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;

            // Formatting and labeling
            outSheet.Name = outSheetName;
            outSheet.Range["A2"].Font.Bold = true;
            outSheet.Range["A2"].Value = outSheetName;
        }
    }
}
