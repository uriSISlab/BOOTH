using Microsoft.Office.Interop.Excel;
using System.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.DS200
{
    class DS200_Importer : LogImporter
    {
        public DS200_Importer() : base(new string[][] {
            new string[] { "Text files", "*.txt" }})
        {
        }

        protected override void ImportFileToSheet(string filePath, Worksheet sheet)
        {
            // importing text file as a query table
            QueryTable queryTable = sheet.QueryTables.Add(Connection: "TEXT;" + filePath,
                Destination: sheet.Range["$A$1"]);
            queryTable.Name = Path.GetFileName(filePath);
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
            queryTable.TextFileCommaDelimiter = true;
            queryTable.TextFileSpaceDelimiter = false;
            XlColumnDataType[] columnDataTypes = {
                    XlColumnDataType.xlTextFormat, XlColumnDataType.xlTextFormat, XlColumnDataType.xlTextFormat,
                    XlColumnDataType.xlTextFormat, XlColumnDataType.xlTextFormat, XlColumnDataType.xlTextFormat,
                    XlColumnDataType.xlTextFormat
                };
            queryTable.TextFileColumnDataTypes = columnDataTypes;
            queryTable.TextFileTrailingMinusNumbers = true;
            queryTable.Refresh(BackgroundQuery: false);

            // Rename the worksheet to the file name of the selected data file
            sheet.Name = Util.Clip("Precinct " + Path.GetFileNameWithoutExtension(filePath), 31);
        }
    }
}
