using BOOTH.LogProcessors;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.IO;

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

        public static void ProcessSheetWithProcessor(Worksheet sheet, ILogProcessor processor)
        {
            Sheets sheets = ThisAddIn.app.ActiveWorkbook.Sheets;

            // Check if the data chosen was already processed
            for (int n = 1; n <= sheets.Count; n++)
            {
                if (sheets[n].Name == Util.GetProcessedName(sheet.Name))
                {
                    return;
                }
            }

            // Disable UI updates
            ThisAddIn.app.ScreenUpdating = false;


            Worksheet outputSheet = AddSheetForOutput(sheet);
            System.Diagnostics.Trace.WriteLine("About to construct FastSheetReader at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            FastSheetReader reader = new FastSheetReader(sheet, processor.GetSeparator());
            System.Diagnostics.Trace.WriteLine("Finished constructing FastSheetReader at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
            FastSheetWriter writer = new FastSheetWriter(outputSheet);
            System.Diagnostics.Trace.WriteLine("Finished constructing FastSheetWriter at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());

            Util.RunPipeline(reader, processor, writer, true);
            System.Diagnostics.Trace.WriteLine("Finished running pipeline at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());

            writer.Flush();
            System.Diagnostics.Trace.WriteLine("Flushed writer at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());

            // Tag the sheet with a machine type mark so we don't have to dig into the cell
            // data to identify machine type when trying to generate summary statistics
            outputSheet.CustomProperties.Add(Util.MACHINE_TYPE_MARK_NAME, processor.GetUniqueTag());

            // Re-enable UI updates
            ThisAddIn.app.ScreenUpdating = true;
            System.Diagnostics.Trace.WriteLine("Returning at " + new DateTimeOffset(DateTime.UtcNow).ToUnixTimeSeconds());
        }

        public static void ProcessSheetForLogType(Worksheet sheet, LogType t)
        {
            ILogProcessor processor = Util.CreateProcessor(t);
            ProcessSheetWithProcessor(sheet, processor);
        }

        public static void ProcessEntireDirectory(LogType t)
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

            // Quit if no files found
            if (files.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("No compatible files found in the selected directory.");
                return;
            }

            // Show progress bar
            ProgressBarForm progress = new ProgressBarForm();
            progress.InitializeAndShow(files.Length - 1);

            FileWriter writer = new FileWriter(outputFileName);

            // Loop to process multiple files consecutively
            for (int i = 0; i < files.Length; i++)
            {
                ILogProcessor processor = Util.CreateProcessor(t);
                FileReader reader = new FileReader(files[i]);
                processor.SetFileName(Path.GetFileName(files[i]));

                Util.RunPipeline(reader, processor, writer, writeHeader: i == 0);

                progress.Step();
            }

            writer.Flush();
            progress.Done();
            System.Windows.Forms.MessageBox.Show("Processed output written to " + outputFileName);
        }
    }
}
