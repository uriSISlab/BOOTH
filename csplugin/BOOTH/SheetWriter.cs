using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
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
        private long columnNum;
        private long startRowOffset;
        private long startColumnOffset;
        private long columns;

        public SheetWriter(Worksheet sheet)
        {
            this.sheet = sheet;
            this.rowNum = 1;
            this.columnNum = 1;
            this.columns = 0;
            this.startRowOffset = 0;
            this.startColumnOffset = 0;
        }

        public SheetWriter(Worksheet sheet, long startRowOffset, long startColumnOffset)
        {
            this.sheet = sheet;
            this.rowNum = 1;
            this.columnNum = 1;
            this.columns = 0;
            this.startRowOffset = startRowOffset;
            this.startColumnOffset = startColumnOffset;
        }
        
        public void Done()
        {
            FormatPretty();
        }

        public void FormatPretty()
        {
            string left = Util.GetColumnLetterFromNumber(1 + this.startColumnOffset);
            string right = Util.GetColumnLetterFromNumber(this.columns + this.startColumnOffset);
            this.sheet.Range[left + "1", right + "1"].Font.Bold = true;
            this.sheet.Range[left + "1", right + "1"].Columns.AutoFit();
        }

        public long GetRowNum()
        {
            return rowNum;
        }

        public void WriteLine(params string[] line)
        {
            this.WriteLineArr(line);
        }

        public void LineBreak()
        {
            this.rowNum++;
            this.columnNum = 1;
        }

        public void Return()
        {
            this.columnNum = 1;
        }

        public bool PreviousLine()
        {
            if (this.rowNum == 1)
            {
                return false;
            }
            this.rowNum--;
            this.columnNum = 1;
            return true;
        }

        public void WriteLineArrWithoutLineBreak(string[] line, FieldType[] fieldTypes = null)
        {
            this.columns = (this.columns < line.Length) ? line.Length : columns;
            for (int c = 0; c < line.Length; c++)
            {
                if (line[c] == null)
                {
                    // Do nothing (skip over the cell) for null values
                    continue;
                }
                string cellAddr = Util.GetColumnLetterFromNumber(c + this.columnNum + this.startColumnOffset) + (rowNum + this.startRowOffset);
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
                        this.sheet.Range[cellAddr].Value = DateTime.Parse(line[c]).ToOADate();
                        this.sheet.Range[cellAddr].NumberFormat = "MM/DD/YYYY hh:mm:ss";
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
            this.columnNum += line.Length;
        }

        public void WriteLineArr(string[] line, FieldType[] fieldTypes = null)
        {
            this.WriteLineArrWithoutLineBreak(line, fieldTypes);
            this.LineBreak(); 
        }
    }
}
