using Microsoft.Office.Interop.Excel;

namespace BOOTH.LogProcessors
{
    public abstract class LogSummarizer
    {

        protected abstract string GetDurationFieldName();

        protected abstract string GetDurationFieldColumn();

        protected abstract string GetEventTypeFieldName();


        public void CreateSummaryFrom(Worksheet sheet)
        {
            Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            // Create sheet for output
            string suffix = ".. Stats";
            string prefix = Util.Clip(sheet.Name, (31 - suffix.Length));
            string outSheetName = prefix + suffix;
            Worksheet outSheet = Util.AddSheet(outSheetName);
            if (outSheet == null)
            {
                Util.MessageBox("Could not create new sheet for summary statistics!");
                return;
            }

            string A = this.GetDurationFieldColumn();
            sheet.Range[A + ":" + A].NumberFormat = "mm:ss";

            // Create Pivot Cache from Source Data
            PivotCache pvtCache = workbook.PivotCaches().Create(SourceType: XlPivotTableSourceType.xlDatabase,
                //SourceData: srcData);
                SourceData: sheet.UsedRange);

            // Create Pivot table from Pivot Cache
            //PivotTable pvt = pvtCache.CreatePivotTable(TableDestination: startPvt, TableName: "PivotTable2");
            PivotTable pvt = pvtCache.CreatePivotTable(TableDestination: outSheet.Range["A3"], TableName: "PivotTable2");

            string durationField = this.GetDurationFieldName();
            string eventTypeField = this.GetEventTypeFieldName();

            pvt.AddDataField(pvt.PivotFields(eventTypeField), "Count of " + eventTypeField, XlConsolidationFunction.xlCount);
            pvt.AddDataField(pvt.PivotFields(eventTypeField), "Percent of " + eventTypeField, XlConsolidationFunction.xlCount);
            pvt.PivotFields("Percent of " + eventTypeField).Calculation = XlPivotFieldCalculation.xlPercentOfColumn;
            pvt.AddDataField(pvt.PivotFields(durationField), "Average Duration of " + eventTypeField, XlConsolidationFunction.xlAverage);
            pvt.AddDataField(pvt.PivotFields(durationField), "Max Duration of " + eventTypeField, XlConsolidationFunction.xlMax);
            pvt.AddDataField(pvt.PivotFields(durationField), "Standard Deviation of " + eventTypeField, XlConsolidationFunction.xlStDev);
            pvt.PivotFields("Average Duration of " + eventTypeField).NumberFormat = "mm:ss";
            pvt.PivotFields("Max Duration of " + eventTypeField).NumberFormat = "mm:ss";
            pvt.PivotFields("Standard Deviation of " + eventTypeField).NumberFormat = "mm:ss";

            pvt.PivotFields(eventTypeField).Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField;

            // Formatting and labeling
            outSheet.Name = outSheetName;
            outSheet.Range["A2"].Font.Bold = true;
            outSheet.Range["A2"].Value = outSheetName;
        }
    }
}
