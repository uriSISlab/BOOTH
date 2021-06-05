using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.Dominion_ICE
{
    // Log processing for Dominion ImageCast Evolution Ballot Scanning and Marking Device
    class DICE_Importer : LogProcessors.LogImporter
    {
        public DICE_Importer() : base(new string[][] { new string[] { "Text files", "*.txt" }})
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
                if (lineStr.Length < 21) continue;
                string[] lineArr = new string[2];
                lineArr[0] = lineStr.Substring(0, 20);  // Timestamp is in the first 20 characters
                lineArr[1] = lineStr.Substring(21);
                writer.WriteLineArr(lineArr);
            }
            inputStream.Close();
        }
    }
}
