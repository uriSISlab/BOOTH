using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace BOOTH
{
    static class VSAP_BMD
    {
        public static void Import_VSAPBMD_data()
        {
            
            FileDialog fileDialog = ThisAddIn.app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            // When file explorer opens, only display text log files
            fileDialog.Filters.Clear();
            fileDialog.Filters.Add("Log files", "*.log");
            // OPen the file explorer and allow selection of multiple files
            fileDialog.AllowMultiSelect = true;
            fileDialog.Show();

            // Prevent showing Excel document updates to improve performance
            ThisAddIn.app.ScreenUpdating = false;

            for (int i = 1; i <= fileDialog.SelectedItems.Count; i++)
            {
                // Add an additional sheet and activate it to populate it with VSAP BMD data
                workbook.Sheets.Add(After: workbook.ActiveSheet);
                Worksheet sheet = workbook.ActiveSheet;

                // pulling file path for a specific file
                string filePath = fileDialog.SelectedItems.Item(i);

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
                queryTable.TextFileTabDelimiter = false;
                queryTable.TextFileSemicolonDelimiter = false;
                queryTable.TextFileCommaDelimiter = false;
                queryTable.TextFileSpaceDelimiter = false;
                queryTable.TextFileOtherDelimiter = "|";
                XlColumnDataType[] columnDataTypes = {
                    XlColumnDataType.xlTextFormat, XlColumnDataType.xlGeneralFormat, XlColumnDataType.xlTextFormat,
                    XlColumnDataType.xlTextFormat, XlColumnDataType.xlTextFormat, XlColumnDataType.xlTextFormat,
                    XlColumnDataType.xlTextFormat
                };
                queryTable.TextFileColumnDataTypes = columnDataTypes;
                queryTable.TextFileTrailingMinusNumbers = true;
                queryTable.Refresh(BackgroundQuery: false);

                // Rename the worksheet to the file name of the selected data file
                // TODO: check if name is already taken
                string[] parts = filePath.Split('\\');
                sheet.Name = parts[parts.Length - 1];
            }

            // Allow the excel file to actively update
            ThisAddIn.app.ScreenUpdating = true;
        }
    }
}
