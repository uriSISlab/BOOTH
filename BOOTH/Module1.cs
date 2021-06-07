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
        public static void PollpadImport()
        {
            FileDialog fileDialog = ThisAddIn.app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            // When file explorer opens, only display text log files
            fileDialog.Filters.Clear();
            fileDialog.Filters.Add("PollPad files", "*.txt; *.csv");
            // OPen the file explorer and allow selection of multiple files
            fileDialog.AllowMultiSelect = true;
            fileDialog.Show();

            // Prevent showing Excel document updates to improve performance
            ThisAddIn.app.ScreenUpdating = false;

            // Loop to process multiple files consecutively
            for (int i = 1; i <= fileDialog.SelectedItems.Count; i++)
            {
                // pulling file path for a specific file
                string filePath = fileDialog.SelectedItems.Item(i);
                string fileNameOnly = Util.Clip(Path.GetFileNameWithoutExtension(filePath), 10);

                // Check for duplicate precincts and delete the duplicate sheets
                int c = 1;
                bool skip = false;
                while (c < workbook.Sheets.Count + 1)
                {
                    if (workbook.Sheets[c].name == fileNameOnly + " PollPad")
                    {
                        Util.MessageBox(fileNameOnly + " shares the first 10 characters with a current worksheet."
                            + " Please rename the file and import again.");
                        break;
                    }
                    c++;
                }
                if (skip)
                {
                    continue;
                }

                // Add an additional sheet and activate it to populate it with DS200 data
                workbook.Sheets.Add(After: workbook.ActiveSheet);
                Worksheet sheet = workbook.ActiveSheet;

                // importing text file as a query table
                QueryTable queryTable = sheet.QueryTables.Add(Connection: "TEXT;" + filePath,
                    Destination: sheet.Range["$A$1"]);
                queryTable.Name = "Precinct " + i;
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
                queryTable.TextFileTrailingMinusNumbers = true;
                queryTable.Refresh(BackgroundQuery: false);

                // Rename the worksheet to the file name of the selected data file
                sheet.Name = fileNameOnly + " PollPad";
            }

            // Allow the excel file to actively update
            ThisAddIn.app.ScreenUpdating = true;
        }

        public static void PollPadProcessing()
        {
            ThisAddIn.app.ScreenUpdating = false;
            Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            Worksheet sheet = workbook.ActiveSheet;
            int colNum = sheet.UsedRange.Columns.Count;

            // Loops through worksheet to format data, separating date and time
            for (int i = 1; i <= colNum; i++)
            {
                if (sheet.Cells[2, i].NumberFormat == "m/d/yyyy h:mm")
                {
                    sheet.Columns[i + 1].Insert();
                    sheet.Columns[i].Copy(sheet.Columns[i + 1]);
                    sheet.Columns[i + 1].NumberFormat = "h:mm";
                    sheet.Cells[1, i + 1] = "Time";
                    sheet.Columns[i + 1].Insert();
                    sheet.Columns[i].Copy(sheet.Columns[i + 1]);
                    sheet.Columns[i + 1].NumberFormat = "m/d/yyyy";
                    sheet.Cells[1, i + 1] = "Date";
                    break;
                }
            }
            ThisAddIn.app.ScreenUpdating = true;
        }

        public static void PivotTablePollPad()
        {
            ThisAddIn.app.ScreenUpdating = false;
            Workbook activeWorkbook = ThisAddIn.app.ActiveWorkbook;
            Worksheet activeSheet = ThisAddIn.app.ActiveWorkbook.ActiveSheet;

            // Storing shortened file names
            string firstName = activeSheet.Name;
            string secondName = firstName.Substring(0, 10) + " PrecinctTurnout";
            string thirdName = firstName.Substring(0, 10) + " TotalTurnout";
            long rawRows = activeSheet.Cells[activeSheet.Rows.Count, 1].End(XlDirection.xlUp).Row;

            int i = 0;
            bool skip = false;

            // Tests to see if sheet name is already taken
            for (int y = 1; y <= activeWorkbook.Sheets.Count; y++)
            {
                Util.MessageBox("Sheet name already taken for precinct turnout, please rename the sheet.");
                i = 1;
                skip = true;
            }

            if (!skip)
            {
                // Filters the PollPad data by time in ascending order
                activeSheet.AutoFilterMode = false;

                activeSheet.Range["C1"].Select();
                ThisAddIn.app.Selection.AutoFilter();
                activeSheet.AutoFilter.Sort.SortFields.Clear();
                activeSheet.AutoFilter.Sort.SortFields.Add(Key: activeSheet.Range["D:D"],
                    SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending,
                    DataOption: XlSortDataOption.xlSortNormal);
                activeSheet.AutoFilter.Sort.Header = XlYesNoGuess.xlGuess;
                activeSheet.AutoFilter.Sort.MatchCase = false;
                activeSheet.AutoFilter.Sort.Orientation = XlSortOrientation.xlSortRows;
                activeSheet.AutoFilter.Sort.SortMethod = XlSortMethod.xlPinYin;
                activeSheet.AutoFilter.Sort.Apply();

                // Sets the starting point of the dayan hour before the first observation
            }
        }
    }
}
