using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;

namespace BOOTH
{
    static class Dispatch
    {
        private static Worksheet AddSheetForOutput(Worksheet afterSheet)
        {
            ThisAddIn.app.ActiveWorkbook.Worksheets.Add(After: afterSheet);
            ThisAddIn.app.ActiveWorkbook.ActiveSheet.Name = Util.GetProcessedName(afterSheet.Name);
            return ThisAddIn.app.ActiveWorkbook.ActiveSheet;
        }

        public static void ProcessSheetForLogType(Worksheet sheet, LogType t)
        {

            Sheets sheets = ThisAddIn.app.ActiveWorkbook.Sheets;
            ILogProcessor processor = Util.CreateProcessor(t);

            // Check if the data chosen was already processed
            for (int n = 1; n <= sheets.Count; n++)
            {
                if (sheets[n].Name == Util.GetProcessedName(sheet.Name))
                {
                    return;
                }
            }
            
            SheetReader reader = new SheetReader(sheet, processor.GetSeparator());
            SheetWriter writer = new SheetWriter(AddSheetForOutput(sheet));

            Util.RunPipeline(reader, processor, writer, true);

            writer.FormatPretty();
        }

        public static void processEntireDirectory(LogType t)
        {
            String folder;
            // Create folder picker
            FileDialog fileDialog = ThisAddIn.app.FileDialog[MsoFileDialogType.msoFileDialogFolderPicker];
            fileDialog.AllowMultiSelect = false;
            if (fileDialog.Show() != -1)
            {
                return;
            }
            folder = fileDialog.SelectedItems.Item(1);

            // Prevent showing excel document updates to improve performance
            ThisAddIn.app.ScreenUpdating = false;

            string outputFileName = Path.Combine(folder, "processed_all.csv");
            string[] files = Directory.GetFiles(folder, Util.GetFileNamePatternForLog(t));

            // Show progress bar
            // TODO Port UserForm1
            // progress 0

            FileWriter writer = new FileWriter(outputFileName);

            // Loop to process multiple files consecutively
            for (int i = 0; i < files.Length; i++)
            {
                ILogProcessor processor = Util.CreateProcessor(t);
                FileReader reader = new FileReader(files[i]);
                processor.SetFileName(Path.GetFileName(files[i]));

                Util.RunPipeline(reader, processor, writer, writeHeader: i == 0);

                // If fileNum Mod 5 = 0 Then
                // progress fileNum / fileCount * 100
                // End If
            }

            writer.Done();

            // Stop showing progress bar
            // Unload UserForm1
        }
    }
}
