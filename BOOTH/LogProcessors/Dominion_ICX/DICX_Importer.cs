﻿using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;

namespace BOOTH.LogProcessors.Dominion_ICX
{
    // Log processing for Dominion ImageCast X Ballot Scanning and Marking device
    class DICX_Importer : LogImporter
    {
        public DICX_Importer() : base(new string[][] { new string[] { "Log files", "*.log" } })
        {
        }

        protected override bool IsCorrectLogType(string filePath)
        {
            // Open the file as a text stream for reading
            StreamReader inputStream = new StreamReader(filePath);
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                if (lineStr.Contains("- LIFETIME COUNTER: "))
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
            string[,] table = new string[numLines, 2];
            uint line = 0;
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                // TODO test if line is well-formed (has a timestamp)
                if (lineStr.Length < 23) continue;
                table[line, 0] = lineStr.Substring(0, 19);  // Timestamp is in the first 19 characters
                table[line, 1] = lineStr.Substring(22);     // Next three characters are " - ", so the rest of the line starts from 22.
                line += 1;
            }
            inputStream.Close();
            Range topLeft = sheet.Cells[1, 1];
            Range bottomRight = sheet.Cells[numLines, 2];
            sheet.Range[topLeft, bottomRight].Value2 = table;
        }

    }
}
