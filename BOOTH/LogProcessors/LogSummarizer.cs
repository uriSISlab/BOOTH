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

        protected virtual ColumnInfo[] GetCategoricalColumnInfos()
        {
            return new ColumnInfo[] { };
        }

        public void CreateSummaryFrom(Worksheet sheet)
        {
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

            System.Diagnostics.Trace.WriteLine("Created new sheet at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            this.CreateMainPivotTable(sheet, outSheet);
            System.Diagnostics.Trace.WriteLine("Created main pivot table at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            ColumnInfo timestampColumn = this.GetTimestampColumnInfo();
            Range timestampRange = sheet.Range[String.Format("{0}2:{0}50", timestampColumn.columnId)];
            // AddTimeStampContinuousChart(sheet, outSheet, timestampRange);
            CreateCategoricalPieCharts(sheet, outSheet);
            System.Diagnostics.Trace.WriteLine("Created categorical pie chart(s) at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            object[,] timestamps = GetColumn(sheet, timestampColumn);
            Range outRange = AddTimestampSummaryTable(timestamps, outSheet);
            System.Diagnostics.Trace.WriteLine("Created timestamp summary table at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            AddTimeStampChart(outSheet, outRange);
            System.Diagnostics.Trace.WriteLine("Created timestamp chart at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());


            // Formatting and labeling
            // TODO check if this name is already taken
            outSheet.Name = outSheetName;
            outSheet.Range["A2"].Font.Bold = true;
            outSheet.Range["A2"].Value = "Event Statistics";
        }

        private void CreateMainPivotTable(Worksheet inSheet, Worksheet outSheet)
        {
            Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            ColumnInfo durationColumn = this.GetDurationColumnInfo();
            ColumnInfo eventTypeColumn = this.GetEventTypeColumnInfo();
            string A = durationColumn.columnId;
            inSheet.Range[A + ":" + A].NumberFormat = "mm:ss";

            // Create Pivot Cache from Source Data
            PivotCache pvtCache = workbook.PivotCaches().Create(SourceType: XlPivotTableSourceType.xlDatabase,
                //SourceData: srcData);
                SourceData: inSheet.UsedRange);

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
        }

        private void AddTimeStampChart(Worksheet outSheet, Range sourceRange)
        {
            // Add timestamp chart
            var charts = outSheet.ChartObjects(Type.Missing) as ChartObjects;
            Range targetRange = outSheet.Range["K1"];
            var chartObject = charts.Add(targetRange.Left, targetRange.Top, 300, 300) as ChartObject;
            outSheet.Activate();
            sourceRange.Select();
            // chartObject.Select();
            var chart = chartObject.Chart;
            chart.SetSourceData(sourceRange, System.Reflection.Missing.Value);

            chart.ChartType = XlChartType.xlColumnClustered;
            chart.ChartWizard(Title: "Event Count by Hour");
        }

        private void AddTimeStampContinuousChart(Worksheet inSheet, Worksheet outSheet, Range sourceRange)
        {
            // Add timestamp chart
            var charts = outSheet.ChartObjects(Type.Missing) as ChartObjects;
            Range targetRange = outSheet.Range["A20"];
            var chartObject = charts.Add(targetRange.Left, targetRange.Top, 300, 300) as ChartObject;
            sourceRange.Select();
            var chart = chartObject.Chart;
            chart.SetSourceData(sourceRange, System.Reflection.Missing.Value);

            chart.ChartType = (XlChartType)118;
            chart.ChartWizard(Title: "Event time histogram");
            outSheet.Activate();
        }

        private Range AddTimestampSummaryTable(object[,] timestamps, Worksheet outSheet)
        {
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
            countsTable[0, 0] = "Hour (Across all days)";
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
            outSheet.Range["H1:I1"].Font.Bold = true;
            Range outRange = outSheet.Range["H1", "I" + countsTable.GetLength(0)];
            outRange.Value = countsTable;
            outSheet.Columns["H:I"].AutoFit();
            return outSheet.get_Range("$H$1", "$I$" + countsTable.GetLength(0));
        }

        private object[,] GetColumn(Worksheet sheet, ColumnInfo column)
        {
            return sheet.Range[column.columnId + ":" + column.columnId].Value2;
        }

        private void CreateCategoricalPieCharts(Worksheet inSheet, Worksheet outSheet)
        {
            int outColumnIndex = 1;
            foreach (ColumnInfo cinfo in this.GetCategoricalColumnInfos())
            {
                var rows = inSheet.UsedRange.Rows.Count;
                var sourceRange = inSheet.Range[String.Format("{0}1:{0}{1}", cinfo.columnId, rows)];
                Range targetRange = outSheet.Range[Util.GetColumnLetterFromNumber(outColumnIndex) + "27"];
                var pivotCache = ThisAddIn.app.ActiveWorkbook.PivotCaches().Create(XlPivotTableSourceType.xlDatabase, sourceRange,
                    XlPivotTableVersionList.xlPivotTableVersion12);
                var pivotTable = pivotCache.CreatePivotTable(targetRange, "PivotTable" + outColumnIndex);
                pivotTable.AddDataField(pivotTable.PivotFields(cinfo.name), "Count", XlConsolidationFunction.xlCount);
                pivotTable.PivotFields(cinfo.name).Orientation = XlPivotFieldOrientation.xlRowField;

                var charts = outSheet.ChartObjects(Type.Missing) as ChartObjects;
                targetRange = outSheet.Range[Util.GetColumnLetterFromNumber(outColumnIndex + 2) + "27"];
                Range bottomRight = outSheet.Range[Util.GetColumnLetterFromNumber(outColumnIndex + 5) + "37"];
                int width = (int) (bottomRight.Left - targetRange.Left);
                int height = (int) (bottomRight.Top - targetRange.Top);
                var chartObject = charts.Add(targetRange.Left, targetRange.Top, width, height) as ChartObject;
                var chart = chartObject.Chart;
                chart.SetSourceData(pivotTable.TableRange1, System.Reflection.Missing.Value);
                chart.ChartType = XlChartType.xlPie;
                chart.ChartWizard(Title: cinfo.name);
                outColumnIndex += 6;
            }
        }
    }
}
