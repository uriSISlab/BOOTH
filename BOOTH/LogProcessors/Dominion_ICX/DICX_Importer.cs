using BOOTH.LogProcessors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.Dominion_ICX
{
    // Log processing for Dominion ImageCast X Ballot Scanning and Marking device
    class DICX_Importer: LogImporter
    {
        public DICX_Importer() : base(new string[][] { new string[] { "Log files", "*.log" } })
        {
        }

        protected override void ImportFileToSheet(string filePath, Worksheet sheet)
        {
            // Open the file as a text stream for reading
            StreamReader inputStream = new StreamReader(filePath);
            SheetWriter writer = new SheetWriter(sheet);
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                // TODO test if line is well-formed (has a timestamp)
                if (lineStr.Length < 23) continue;
                string[] lineArr = new string[2];
                lineArr[0] = lineStr.Substring(0, 19);  // Timestamp is in the first 19 characters
                lineArr[1] = lineStr.Substring(22);     // Next three characters are " - ", so the rest of the line starts from 22.
                writer.WriteLineArr(lineArr);
            }
            inputStream.Close();
        }

    }
}
