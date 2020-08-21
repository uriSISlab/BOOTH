using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    }
}
