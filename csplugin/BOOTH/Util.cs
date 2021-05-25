using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOOTH
{
    public enum LogType
    {
        VSAP_BMD,
        DICE,
        DICX,
        DS200,
        POLLPAD,
        UNKNOWN
    }

    public enum IOType
    {
        FILE,
        SHEET
    }

    public enum FieldType
    {
        INTEGER,
        FLOATING,
        TIMESPAN_MMSS,
        DATETIME,
        STRING
    }

    public static class Util
    {
        public static string GetColumnLetterFromNumber(long number)
        {
            return GetColumnLetterFromNumberZeroBased(number - 1);
        }

        public static string GetColumnLetterFromNumberZeroBased(long number)
        {
            string lastLetter = ((char)('A' + (number % 26))).ToString();
            return (number >= 26) ? (GetColumnLetterFromNumber(number / 26) + lastLetter) : lastLetter;
        }

        public static int GetDifferenceMinutes(DateTime startTime, DateTime endTime)
        {
            return (int)((endTime - startTime).TotalSeconds) / 60;
        }

        public static string GetTimeDifference(DateTime startTime, DateTime endTime)
        {
            return (endTime - startTime).ToString(@"mm\:ss");
        }

        public static string GetProcessedName(string name)
        {
            string processedName = name + " Processed";
            if (processedName.Length > 31)
            {
                processedName = processedName.Substring(processedName.Length - 31, 31);
            }
            return processedName;
        }

        public static ILogProcessor CreateProcessor(LogType t)
        {
            switch (t)
            {
                case LogType.VSAP_BMD:
                    return new VSAPBMD_Processor();
                case LogType.DICE:
                    return new DICE_Processor();
                case LogType.DICX:
                    return new DICX_Processor();
                case LogType.DS200:
                    return new DS200_Processor();
                default:
                    return null;
            }
        }

        public static string GetFileNamePatternForLog(LogType t)
        {
            switch (t)
            {
                case LogType.VSAP_BMD:
                    return "BEL_*_*.log";
                case LogType.DICE:
                    return "*.TXT";
                case LogType.DICX:
                    return "ICX_AUDIT_LOG.*.log";
                case LogType.DS200:
                    return "*.TXT";
                default:
                    return "*.*";
            }
        }

        public static ILogProcessor GetCorrectProcessorForImportedSheet(Worksheet sheet)
        {
            // Go through our log processors and return the correct one for this sheet 
            foreach (LogType logType in Enum.GetValues(typeof(LogType)))
            {
                ILogProcessor processor = CreateProcessor(logType);
                if (processor != null && processor.IsThisLog(sheet))
                {
                    return processor;
                }
            }
            return null;
        }

        public static string Clip(string input, int length)
        {
            return input.Length > length ? input.Substring(0, length) : input;
        }

        public static void RunPipeline(IInputReader reader, ILogProcessor processor, IOutputWriter writer, bool writeHeader)
        {
            processor.SetWriter(writer);
            if (writeHeader)
            {
                processor.WriteHeader();
            }
            while (!reader.NoMoreLines())
            {
                processor.ReadLine(reader.ReadLine());
            }
            processor.Done();
        }

        public static string[] AppendToArray(string[] arr, string toAppend)
        {
            string[] fullArr;
            fullArr = new string[arr.Length + 1];
            // Copy contents of arr to fullArr first
            System.Array.Copy(arr, 0, fullArr, 0, arr.Length);
            fullArr[fullArr.Length - 1] = toAppend;
            return fullArr;
        }

        public static void MessageBox(string message)
        {
            System.Windows.Forms.MessageBox.Show(message);
        }

        public static Worksheet TryAddingSheetWithName(string name)
        {
            try
            {
                if (ThisAddIn.app.ActiveWorkbook == null)
                {
                    ThisAddIn.app.Workbooks.Add();
                }
                Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets.Add();
                sheet.Name = name;
                return sheet;
            } catch (Exception)
            {
                return null;
            }
        }
    }
}
