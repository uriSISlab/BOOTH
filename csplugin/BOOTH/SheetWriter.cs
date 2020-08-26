using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    public class SheetWriter : IOutputWriter
    {
        private readonly Worksheet sheet;
        private long rowNum;

        public SheetWriter(Worksheet sheet)
        {
            this.sheet = sheet;
            this.rowNum = 1;
        }

        public void Done()
        {
            FormatPretty();
        }

        public void FormatPretty()
        {
            int columns = this.sheet.UsedRange.Columns.Count;
            this.sheet.Range["A1", Util.GetLetterFromNumber(columns) + "1"].Font.Bold = true;
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
            for (int c = 0; c < line.Length; c++)
            {
                string cellAddr = Util.GetLetterFromNumber(c + 1) + rowNum;
                if (fieldTypes == null || fieldTypes.Length - 1 < c)
                {
                    this.sheet.Range[cellAddr].Value = line[c];
                    continue;
                }
                switch (fieldTypes[c])
                {
                    case FieldType.INTEGER:
                        this.sheet.Range[cellAddr].Value = int.Parse(line[c]);
                        break;
                    case FieldType.FLOATING:
                        this.sheet.Range[cellAddr].Value = double.Parse(line[c]);
                        break;
                    case FieldType.DATETIME:
                        // TODO assign properly
                        this.sheet.Range[cellAddr].Value = line[c];
                        break;
                    case FieldType.TIMESPAN_MMSS:
                        this.sheet.Range[cellAddr].Value = TimeSpan.ParseExact(line[c], @"mm\:ss", CultureInfo.InvariantCulture).TotalDays;
                        this.sheet.Range[cellAddr].NumberFormat = "mm:ss";
                        break;
                    case FieldType.STRING:
                    default:
                        this.sheet.Range[cellAddr].Value = line[c];
                        break;
                }
            }  
            rowNum++;
        }
    }
}
