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
        private List<object[]> lines = new List<object[]>();
        private List<string[]> numberFormats = new List<string[]>();

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
            this.GetOutputRange().NumberFormat = Util.JaggedTo2DArray(this.numberFormats.ToArray(), this.columns);
            this.FormatPretty();
        }

        public void FormatPretty()
        {
            string left = Util.GetColumnLetterFromNumber(1 + this.startColumnOffset);
            string right = Util.GetColumnLetterFromNumber(this.columns + this.startColumnOffset);
            this.sheet.Range[left + "1", right + "1"].Font.Bold = true;
            this.sheet.Range[left + "1", right + (this.rowNum - 1)].Columns.AutoFit();
        }

        public long GetRowNum()
        {
            return rowNum;
        }

        public void WriteLine(params string[] line)
        {
            this.WriteLineArr(line);
        }

        public void WriteLineArr(IEnumerable<string> line, IEnumerable<FieldType> fieldTypes = null)
        {
            string[] lineArr = line.ToArray();
            FieldType[] fieldTypesArr = fieldTypes?.ToArray();
            this.columns = Math.Max(this.columns, lineArr.Length);
            object[] lineObjArr = new object[lineArr.Length];
            string[] numFormatArr = new string[lineArr.Length];
            for (int c = 0; c < lineArr.Length; c++)
            {
                if (fieldTypesArr == null || fieldTypesArr.Length - 1 < c)
                {
                    numFormatArr[c] = "@";
                    lineObjArr[c] = lineArr[c];
                    continue;
                }
                switch (fieldTypesArr[c])
                {
                    case FieldType.INTEGER:
                        lineObjArr[c] = int.Parse(lineArr[c]);
                        numFormatArr[c] = "@";
                        break;
                    case FieldType.FLOATING:
                        lineObjArr[c] = double.Parse(lineArr[c]);
                        numFormatArr[c] = "@";
                        break;
                    case FieldType.DATETIME:
                        lineObjArr[c] = DateTime.Parse(lineArr[c]).ToOADate();
                        numFormatArr[c] = "MM/DD/YYYY hh:mm:ss";
                        break;
                    case FieldType.TIMESPAN_MMSS:
                        lineObjArr[c] = TimeSpan.ParseExact(lineArr[c], @"mm\:ss", CultureInfo.InvariantCulture).TotalDays;
                        numFormatArr[c] = "mm:ss";
                        break;
                    case FieldType.STRING:
                    default:
                        numFormatArr[c] = "@";
                        lineObjArr[c] = lineArr[c];
                        break;
                }
            }
            this.lines.Add(lineObjArr);
            this.numberFormats.Add(numFormatArr);
            this.rowNum++;
        }
    }
}
