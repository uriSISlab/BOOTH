using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;
using System;

namespace BOOTH.LogProcessors.VSAP_BMD
{
    class VSAPBMD_Importer : LogImporter
    {
        public VSAPBMD_Importer() : base(new string[][] { new string[] { "Log files", "*.log" } })
        {
        }

        protected override bool IsCorrectLogType(string filePath)
        {
            // Open the file as a text stream for reading
            StreamReader inputStream = new StreamReader(filePath);
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                if (lineStr.Contains("|Logger.js-Loading page-Manual Diagnostic Status|system|info|"))
                {
                    return true;
                }
            }
            return false;
        }

        protected override void ImportFileToSheet(string filePath, Worksheet sheet)
        {
            // Open the file as a text stream for reading
            StreamReader inputStream = new StreamReader(filePath);
            int numLines = File.ReadLines(filePath).Count();
            string[][] table = new string[numLines][];
            int line = 0;
            int maxCols = 0;
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                // The pipe character is used to separate fields in
                // VSAP BMD logs.
                string[] lineArr = lineStr.Split('|');
                table[line] = lineArr;
                maxCols = Math.Max(maxCols, lineArr.Length);
                line += 1;
            }
            inputStream.Close();
            Range topLeft = sheet.Cells[1, 1];
            Range bottomRight = sheet.Cells[numLines, maxCols];
            sheet.Range[topLeft, bottomRight].Value2 = Util.JaggedTo2DArray(table, maxCols);
        }
    }
}
