using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    // Log processing for Dominion ImageCast X Ballot Scanning and Marking device
    static class Dominion_ICX
    {
        public static void Import_DICX_Data()
        {
            // When File Explorer opens, only display text log files
            FileDialog fileDialog = ThisAddIn.app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            fileDialog.Filters.Clear();
            fileDialog.Filters.Add("Log files", "*.log");
            // Open the file explorer and allow the selection of multiple files
            fileDialog.AllowMultiSelect = true;
            fileDialog.Show();

            // Prevent showing Excel document updates to improve performance
            ThisAddIn.app.ScreenUpdating = false;

            Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            // Loop to process multiple files consecutively
            for (int j = 1; j <= fileDialog.SelectedItems.Count; j++)
            {
                // Add an additional sheet and activate it to populate it with Dominion ICE data
                workbook.Sheets.Add(After: workbook.ActiveSheet);

                // Pulling file path for a specific file
                string filePath = fileDialog.SelectedItems.Item(j);

                Import_DICX_File_Into_Sheet(filePath, workbook.ActiveSheet);

                // Rename the Worksheet to the file name of the selected data file
                // TODO: check if name is already taken
                string[] parts = filePath.Split('\\');
                workbook.ActiveSheet.Name = parts[parts.Length - 1];
            }

            // Allow the Excel file to actively update
            ThisAddIn.app.ScreenUpdating = true;
        }
        public static void Import_DICX_File_Into_Sheet(string filePath, Worksheet sheet)
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
                lineArr[1] = lineStr.Substring(21);     // Next three characters are " - ", so the rest of the line starts from 23.
                writer.WriteLineArr(lineArr);
            }
            inputStream.Close();
        }

    }
}
