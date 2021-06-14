using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.PollPad
{
    class PollPad_Importer : LogImporter
    {
        private static readonly string monthPattern = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)";
        private static readonly string timestampPattern = @"\A" + monthPattern + @" \d\d?, \d\d\d\d at " + @"\d\d?:\d\d:\d\d (AM|PM)";
        public static readonly Regex timestampRegex = new Regex(timestampPattern, RegexOptions.Compiled);

        public PollPad_Importer() : base (new string[][] { new string[] { "Text files", "*.txt"}})
        {
        }

        protected override bool IsCorrectLogType(string filePath)
        {
            // Open the file as a text stream for reading
            StreamReader inputStream = new StreamReader(filePath);
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                if (lineStr.Contains("| CHECK IN VOTER |"))
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
                string[] lineArray = lineArr.Select(s => s.Trim()).ToArray();
                table[line] = lineArray;
                maxCols = Math.Max(maxCols, lineArray.Length);
                line += 1;
            }
            inputStream.Close();
            Range topLeft = sheet.Cells[1, 1];
            Range bottomRight = sheet.Cells[numLines, maxCols];
            sheet.Range[topLeft, bottomRight].Value2 = Util.JaggedTo2DArray(table, maxCols);
        }
    }
}
