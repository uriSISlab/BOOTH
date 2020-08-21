using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    class SheetReader : IInputReader
    {

        private readonly Worksheet sheet;
        private readonly string separator;
        private readonly int columns;
        private long currentLine;

        public SheetReader(Worksheet sheet, string separator)
        {
            this.sheet = sheet;
            this.separator = separator;
            this.columns = sheet.UsedRange.Columns.Count;
            this.currentLine = 1;
        }

        public bool NoMoreLines()
        {
            return this.currentLine > this.sheet.UsedRange.Rows.Count;
        }

        public string ReadLine()
        {
            string line = sheet.Range["A" + currentLine].Text;
            for (int j = 2; j <= this.columns; j++)
            {
                // Join the row with the separator
                line = line + separator + sheet.Range[Util.GetLetterFromNumber(j) + this.currentLine].Text.ToString();
            }
            this.currentLine++;
            return line;
        }

        public void SetSkipLines(int skipCount)
        {
            this.currentLine = 1 + skipCount;
        }
    }
}
