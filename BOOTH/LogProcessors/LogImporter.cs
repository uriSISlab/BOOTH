using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors
{
    public abstract class LogImporter
    {
        private readonly string[][] fileTypeFilters;

        public LogImporter(string[][] fileTypeFilters) {
            this.fileTypeFilters = fileTypeFilters;
        }

        protected abstract void ImportFileToSheet(string filePath, Worksheet sheet);

        public void ImportIntoCurrentSheet()
        {
            FileDialog fileDialog = ThisAddIn.app.FileDialog[MsoFileDialogType.msoFileDialogFilePicker];
            fileDialog.Filters.Clear();
            for (int i = 0; i < fileTypeFilters.Length; i++)
            {
                fileDialog.Filters.Add(fileTypeFilters[i][0], fileTypeFilters[i][1]);
            }

            fileDialog.AllowMultiSelect = true;
            fileDialog.Show();

            // Prevent showing Excel document updates to improve performance
            ThisAddIn.app.ScreenUpdating = false;

            Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            // Loop to process multiple files consecutively
            for (int j = 1; j <= fileDialog.SelectedItems.Count; j++)
            {
                // Pull file path for the specific file
                string filePath = fileDialog.SelectedItems.Item(j);

                // Add an additional sheet and activate it
                string sheetName = Path.GetFileNameWithoutExtension(filePath);
                Worksheet addedSheet = Util.AddSheet(sheetName, workbook.ActiveSheet);

                this.ImportFileToSheet(filePath, addedSheet);
                addedSheet.Columns.AutoFit();
            }


            // Allow the Excel file to actively update
            ThisAddIn.app.ScreenUpdating = true;
        }
    }
}
