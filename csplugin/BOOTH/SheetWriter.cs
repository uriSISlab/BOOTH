using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    class SheetWriter : IOutputWriter
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
            this.sheet.Range["A1", Util.getLetterFromNumber(columns) + "1"].Font.Bold = true;
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

        public void WriteLineArr(string[] line)
        {
            string rangeEnd = Util.getLetterFromNumber(line.Length) + rowNum;
            this.sheet.Range["A" + rowNum, rangeEnd].Value = line;
            rowNum++;
        }
    }
}
