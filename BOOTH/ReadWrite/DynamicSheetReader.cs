using Microsoft.Office.Interop.Excel;

namespace BOOTH
{
    class DynamicSheetReader : IInputReader
    {

        private readonly Worksheet sheet;
        private readonly string separator;
        private readonly int columns;
        private long currentLine;

        public DynamicSheetReader(Worksheet sheet, string separator)
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
                line = line + separator + sheet.Range[Util.GetColumnLetterFromNumber(j) + this.currentLine].Text.ToString();
            }
            this.currentLine++;
            return line;
        }
    }
}
