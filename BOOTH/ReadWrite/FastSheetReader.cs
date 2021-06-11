using Microsoft.Office.Interop.Excel;
using System;

namespace BOOTH
{
    class FastSheetReader : IInputReader
    {

        private readonly Worksheet sheet;
        private readonly string separator;
        private readonly int columns;
        private readonly int rows;
        private long currentLine;
        private readonly object[,] data;

        public FastSheetReader(Worksheet sheet, string separator)
        {
            this.sheet = sheet;
            this.separator = separator;
            this.columns = sheet.UsedRange.Columns.Count;
            this.rows = sheet.UsedRange.Rows.Count;
            this.currentLine = 1;
            Range all = sheet.get_Range("A1", Util.GetColumnLetterFromNumber(columns) + rows);
            if (this.rows == 1 && this.columns == 1)
            {
                // If only one cell is non-empty in the sheet, Range.Cells.Value is not an array.
                // We have to use this clunky method to instantiate a 1-indexed array. We need a 1-indexed
                // since all.Cells.Value returns a 1-indexed array when there is more than one cell occupied.
                this.data = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
                this.data[1, 1] = all.Cells.Value;
            } else
            {
                this.data = all.Cells.Value;
            }
        }

        public bool NoMoreLines()
        {
            return this.currentLine > this.rows;
        }

        public string ReadLine()
        {
            string line = this.data[currentLine, 1].ToString();
            for (int j = 2; j <= this.columns; j++)
            {
                object item = this.data[currentLine, j];
                if (item != null)
                {
                    // Join the row with the separator
                    line = line + separator + item.ToString();
                }
            }
            this.currentLine++;
            return line;
        }
    }
}
