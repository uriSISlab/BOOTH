using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.Pollpad
{
    class PollPad_Importer : LogImporter
    {
        private static readonly string monthPattern = "(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)";
        private static readonly string timestampPattern = @"\A" + monthPattern + @" \d\d?, \d\d\d\d at " + @"\d\d?:\d\d:\d\d (AM|PM)";
        private static readonly Regex timestampRegex = new Regex(timestampPattern, RegexOptions.Compiled);

        public PollPad_Importer() : base (new string[][] { new string[] { "Text files", "*.txt"}})
        {
        }

        protected override void ImportFileToSheet(string filePath, Worksheet sheet)
        {
            // Open the file as a text stream for reading
            StreamReader inputStream = new StreamReader(filePath);
            SheetWriter writer = new SheetWriter(sheet);
            System.Diagnostics.Debug.WriteLine("Reading pollpad log from file " + filePath);
            int lines = 0;
            while (!inputStream.EndOfStream)
            {
                string lineStr = inputStream.ReadLine();
                // Test if line is well-formed (has a timestamp)
                if (!timestampRegex.IsMatch(lineStr))
                {
                    continue;
                }
                lines++;
                string[] lineArr = lineStr.Split('|');
                IEnumerable<string> lineArray = lineArr.Select(s => s.Trim());
                writer.WriteLineArr(lineArray);
            }
            inputStream.Close();
            System.Diagnostics.Debug.WriteLine(lines + " lines read.");
        }
    }
}
