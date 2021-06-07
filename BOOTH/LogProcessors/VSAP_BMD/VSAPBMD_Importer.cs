using BOOTH.LogProcessors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace BOOTH.LogProcessors.VSAP_BMD
{
    class VSAPBMD_Importer : LogImporter
    {
        public VSAPBMD_Importer() : base(new string[][] { new string[] {"Log files", "*.log"}})
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
                // The pipe character is used to separate fields in
                // VSAP BMD logs.
                string[] lineArr = lineStr.Split('|');
                writer.WriteLineArr(lineArr);
            }
            inputStream.Close();
        }
    }
}
