using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace BOOTH
{
    public class FastSheetWriter : IOutputWriter
    {
        private readonly Worksheet sheet;
        private readonly long startRowOffset = 0;
        private readonly long startColumnOffset = 0;
        private long columns = 0;
        private long rowNum = 1;
        private readonly List<object[]> lines = new List<object[]>();
        private readonly List<string[]> numberFormats = new List<string[]>();

        public FastSheetWriter(Worksheet sheet)
        {
            this.sheet = sheet;
        }

        public FastSheetWriter(Worksheet sheet, long startRowOffset, long startColumnOffset)
        {
            this.sheet = sheet;
            this.startRowOffset = startRowOffset;
            this.startColumnOffset = startColumnOffset;
        }

        private Range GetOutputRange()
        {
            Range topLeft = this.sheet.Cells[1 + this.startRowOffset, 1 + this.startColumnOffset];
            Range bottomRight = this.sheet.Cells[this.rowNum - 1 + this.startRowOffset,
                this.columns + this.startColumnOffset];
            return sheet.get_Range(topLeft, bottomRight);
        }

        public void Flush()
        {
            this.GetOutputRange().Value = Util.JaggedTo2DArray(this.lines.ToArray(), this.columns);
            System.Diagnostics.Trace.WriteLine("Finished writing values at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            this.GetOutputRange().NumberFormat = Util.JaggedTo2DArray(this.numberFormats.ToArray(), this.columns);
            System.Diagnostics.Trace.WriteLine("Finished writing numformats at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            this.FormatPretty();
            System.Diagnostics.Trace.WriteLine("Finished format pretty at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
        }

        public void FormatPretty()
        {
            // TODO use startRowOffset and startColumnOffset
            this.sheet.Rows[1].Font.Bold = true;
            this.sheet.UsedRange.Columns.AutoFit();
        }

        public long GetRowNum()
        {
            return rowNum;
        }

        public void WriteLine(params string[] line)
        {
            this.WriteLineArr(line);
        }

        public void WriteLineArr(string[] line, FieldType[] fieldTypes = null)
        {
            FieldType[] fieldTypesArr = fieldTypes?.ToArray();
            this.columns = Math.Max(this.columns, line.Length);
            object[] lineObjArr = new object[line.Length];
            string[] numFormatArr = new string[line.Length];
            for (int c = 0; c < line.Length; c++)
            {
                if (fieldTypesArr == null || fieldTypesArr.Length - 1 < c)
                {
                    numFormatArr[c] = "@";
                    lineObjArr[c] = line[c];
                    continue;
                }
                switch (fieldTypesArr[c])
                {
                    case FieldType.INTEGER:
                        lineObjArr[c] = int.Parse(line[c]);
                        numFormatArr[c] = "@";
                        break;
                    case FieldType.FLOATING:
                        lineObjArr[c] = double.Parse(line[c]);
                        numFormatArr[c] = "@";
                        break;
                    case FieldType.DATETIME:
                        lineObjArr[c] = DateTime.Parse(line[c]).ToOADate();
                        numFormatArr[c] = "MM/DD/YYYY hh:mm:ss";
                        break;
                    case FieldType.TIMESPAN_MMSS:
                        lineObjArr[c] = TimeSpan.ParseExact(line[c], @"mm\:ss", CultureInfo.InvariantCulture).TotalDays;
                        numFormatArr[c] = "mm:ss";
                        break;
                    case FieldType.STRING:
                    default:
                        numFormatArr[c] = "@";
                        lineObjArr[c] = line[c];
                        break;
                }
            }
            this.lines.Add(lineObjArr);
            this.numberFormats.Add(numFormatArr);
            this.rowNum++;
        }
    }
}
