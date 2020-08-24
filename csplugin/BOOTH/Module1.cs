using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    public static class Module1
    {
        public static void Import_DS200_data()
        {
            // When File Explorer opens, only display text files
            FileDialog fileDialog = ThisAddIn.app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            fileDialog.Filters.Clear();
            fileDialog.Filters.Add("Text files", "*.txt");

            // Open the file explorer and allow the selection of multiple files
            fileDialog.AllowMultiSelect = true;
            fileDialog.Show();

            // Prevent showing Excel document updates to improve performance
            ThisAddIn.app.ScreenUpdating = false;

            Workbook activeWorkbook = ThisAddIn.app.ActiveWorkbook;
            // Loop to process multiple files consecutively
            for (int j = 1; j <= fileDialog.SelectedItems.Count; j++)
            {
                // Adds an additional Worksheet to write DS200 data to if only one sheet is open
                if (activeWorkbook.Sheets.Count == 1)
                {
                    activeWorkbook.Sheets.Add(After: activeWorkbook.ActiveSheet);
                }

                // Pulling file path for a specific file
                string nam = fileDialog.SelectedItems.Item(j);

                // Check for duplicate precincts and delete the duplicate sheets
                bool continueThis = false;
                for (int c = 1;  c <= activeWorkbook.Sheets.Count; c++)
                {
                    if (activeWorkbook.Sheets[c].Name == "Precinct " + Path.GetFileNameWithoutExtension(nam))
                    {
                        continueThis = true;
                        break;
                    }
                }
                if (continueThis)
                {
                    continue;
                }

                // Add an additional sheet and activate it to populate it with DS200 data
                activeWorkbook.Sheets.Add(After: activeWorkbook.Sheets[j]);
                activeWorkbook.Sheets[j + 1].Activate();

                // Importing text file as a query table
                QueryTable queryTable = activeWorkbook.ActiveSheet.QueryTables.Add(Connection: "TEXT;" + nam,
                    Destination: activeWorkbook.ActiveSheet.Range["$A$1"]);
                queryTable.Name = "Precinct " + j;
                queryTable.FieldNames = true;
                queryTable.RowNumbers = false;
                queryTable.FillAdjacentFormulas = false;
                queryTable.PreserveFormatting = true;
                queryTable.RefreshOnFileOpen = false;
                queryTable.RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells;
                queryTable.SavePassword = false;
                queryTable.SaveData = true;
                queryTable.AdjustColumnWidth = true;
                queryTable.RefreshPeriod = 0;
                queryTable.TextFilePromptOnRefresh = false;
                queryTable.TextFilePlatform = 437;
                queryTable.TextFileStartRow = 1;
                queryTable.TextFileParseType = XlTextParsingType.xlDelimited;
                queryTable.TextFileTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote;
                queryTable.TextFileConsecutiveDelimiter = false;
                queryTable.TextFileTabDelimiter = true;
                queryTable.TextFileSemicolonDelimiter = false;
                queryTable.TextFileCommaDelimiter = true;
                queryTable.TextFileSpaceDelimiter = false;
                XlColumnDataType[] columnDataTypes = {
                    XlColumnDataType.xlGeneralFormat, XlColumnDataType.xlSkipColumn, XlColumnDataType.xlTextFormat,
                    XlColumnDataType.xlSkipColumn, XlColumnDataType.xlSkipColumn, XlColumnDataType.xlTextFormat,
                    XlColumnDataType.xlTextFormat
                };
                queryTable.TextFileColumnDataTypes = columnDataTypes;
                queryTable.TextFileTrailingMinusNumbers = true;
                queryTable.Refresh(BackgroundQuery: false);
                
                // Rename the Worksheet to the file name of the selected data file
                activeWorkbook.ActiveSheet.Name = "Precinct " + Path.GetFileNameWithoutExtension(nam);
            }

            // Deletes any blank sheets while more than one sheet is open
            int d = activeWorkbook.Sheets.Count;
            for (int t = 1; t <= d; t++)
            {
                if (t <= d && t > 1)
                {
                    if (activeWorkbook.Sheets[t].Range("A1").Text.Length == 0)
                    {
                        activeWorkbook.Worksheets[t].Delete();
                        d = activeWorkbook.Sheets.Count;
                        t = 0;
                    }
                }
                d = activeWorkbook.Sheets.Count;
            }

            // Allow the Excel file to actively update
            ThisAddIn.app.ScreenUpdating = true;
        }
    }
}
