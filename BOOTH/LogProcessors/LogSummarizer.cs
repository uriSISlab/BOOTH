using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;

namespace BOOTH.LogProcessors
{
    public abstract class LogSummarizer
    {
        public struct ColumnInfo {
            public readonly string columnId;
            public readonly string name;

            public ColumnInfo(string columnId, string name)
            {
                this.columnId = columnId;
                this.name = name;
            }
        }

        protected abstract ColumnInfo GetDurationColumnInfo();

        protected abstract ColumnInfo GetEventTypeColumnInfo();

        protected virtual ColumnInfo GetTimestampColumnInfo()
        {
            return new ColumnInfo("B", "Timestamp");
        }

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

            Range outRange = AddTimestampSummaryTable(sheet, outSheet);
            AddTimeStampChart(outSheet, outRange);

            ColumnInfo durationColumn = this.GetDurationColumnInfo();
            ColumnInfo eventTypeColumn = this.GetEventTypeColumnInfo();
            string A = durationColumn.columnId;
            sheet.Range[A + ":" + A].NumberFormat = "mm:ss";

            // Create Pivot Cache from Source Data
            PivotCache pvtCache = workbook.PivotCaches().Create(SourceType: XlPivotTableSourceType.xlDatabase,
                //SourceData: srcData);
                SourceData: sheet.UsedRange);

            // Create Pivot table from Pivot Cache
            //PivotTable pvt = pvtCache.CreatePivotTable(TableDestination: startPvt, TableName: "PivotTable2");
            PivotTable pvt = pvtCache.CreatePivotTable(TableDestination: outSheet.Range["A3"], TableName: "PivotTable2");

            string durationField = durationColumn.name;
            string eventTypeField = eventTypeColumn.name;

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
            // TODO check if this name is already taken
            outSheet.Name = outSheetName;
            outSheet.Range["A2"].Font.Bold = true;
            outSheet.Range["A2"].Value = outSheetName;

        }

        private void AddTimeStampChart(Worksheet outSheet, Range outRange)
        {
            // Add timestamp chart
            var charts = outSheet.ChartObjects(Type.Missing) as ChartObjects;
            Range targetRange = outSheet.Range["L1"];
            var chartObject = charts.Add(targetRange.Left, targetRange.Top, 300, 300) as ChartObject;
            outSheet.Activate();
            outRange.Select();
            // chartObject.Select();
            var chart = chartObject.Chart;
            chart.SetSourceData(outRange, System.Reflection.Missing.Value);

            chart.ChartType = XlChartType.xlColumnClustered;
            // chart.ChartWizard(Source: outRange, Title: "Time distribution", CategoryTitle: "Hour", ValueTitle: "Number of Events");
        }

        private Range AddTimestampSummaryTable(Worksheet inSheet, Worksheet outSheet)
        {
            ColumnInfo timestampColumn = this.GetTimestampColumnInfo();
            object[,] timestamps = inSheet.Range[timestampColumn.columnId + ":" + timestampColumn.columnId].Value2;
            Dictionary<int, int> hourCounts = new Dictionary<int, int>();
            int lowestHour = 23;
            int highestHour = 0;
            for (int i = 1; i < timestamps.GetLength(0); i++)
            {
                var item = timestamps[i, 1];
                if (item == null)
                {
                    // Nothing in cell
                    continue;
                }
                if (item != null && item.GetType() == typeof(Double))
                {
                    DateTime timestamp = DateTime.FromOADate((Double)item);
                    if (!hourCounts.ContainsKey(timestamp.Hour))
                    {
                        hourCounts[timestamp.Hour] = 0;
                    }
                    hourCounts[timestamp.Hour] += 1;
                    if (timestamp.Hour < lowestHour)
                    {
                        lowestHour = timestamp.Hour;
                    }
                    if (timestamp.Hour > highestHour)
                    {
                        highestHour = timestamp.Hour;
                    }
                }
            }
            // This throws an exception if lowestHour > highestHour + 2, as happens when the
            // timestamps column doesn't contain any timestamps.
            object[,] countsTable = new object[(highestHour - lowestHour + 1) + 1, 2];
            countsTable[0, 0] = "Hour";
            countsTable[0, 1] = "Number of Events";
            for (int h = lowestHour; h <= highestHour; h++)
            {
                int row = h - lowestHour + 1;
                DateTime time = new DateTime(2000, 01, 01, h, 0, 0);
                //countsTable[row, 0] = String.Format("{0,2}:00-{0,2}:59", h);
                countsTable[row, 0] = time.ToOADate();
                if (hourCounts.ContainsKey(h))
                {
                    countsTable[row, 1] = hourCounts[h];
                } else
                {
                    countsTable[row, 1] = 0;
                }
            }
            outSheet.Range["H:H"].NumberFormat = "hh:mm";
            Range outRange = outSheet.Range["H1", "I" + countsTable.GetLength(0)];
            outRange.Value = countsTable;
            return outSheet.get_Range("$H$1", "$I$" + countsTable.GetLength(0));
        }
    }
}
