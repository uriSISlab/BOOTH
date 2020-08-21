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
        UNKNOWN
    }

    public enum IOType
    {
        FILE,
        SHEET
    }
    public static class Util
    {
        public static string GetLetterFromNumber(int number)
        {
            return ((char)('A' + number - 1)).ToString();
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
                    return null;
                    //return new DICX_Processor();
                default:
                    return null;
            }
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

    }
}
