using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    class DS200_Processor : ILogProcessor
    {

        private const int votingStartedCode =   1004115;
        private const int blankBallotCode =     1004113;
        private const int overvotedBallotCode = 1004111;
        private const int votingCompleteCode =  1004022;
        private const int ballotJamCode =       3013004;
        private const int jamClearedCode =      1004328;
        private const int shutDownCode =        1004016;
        private const int unknownCode =         1004163;
        private const int votingModeCode =      1004056;

        private int[] recognizedCodes = new int[]
        {
            1004115, 1004163, 3013006, 1004138, 1004016, 1004056,
            1004022, 1004111, 1004113, 3013005, 3003337, 3013001,
            3013004, 3013008, 3013002, 7003009, 3013003, 3013007,
            3013009, 3003335, 3003336, 3003339, 3003340, 3003318,
            3003341, 1004122, 1004112, 1004114, 1004328
        };

        private FieldType[] fieldTypes = new FieldType[] { FieldType.TIMESPAN_MMSS, FieldType.STRING,
            FieldType.STRING, FieldType.INTEGER};

        private string fileName;
        private IOutputWriter writer;
        private readonly List<string> lines;

        public DS200_Processor()
        {
            this.fileName = "";
            this.lines = new List<string>();
        }

        public string GetSeparator()
        {
            return ", ";
        }

        public bool IsThisLog(Worksheet sheet)
        {
            return sheet.Range["A1"].Text.ToString().Trim() == "1114111";
        }

        public void ReadLine(string line)
        {
            if (recognizedCodes.Contains(GetCode(line)))
            {
                this.lines.Add(line);
            }
        }

        public void SetFileName(string fileName)
        {
            this.fileName = fileName;
        }

        public void SetWriter(IOutputWriter writer)
        {
            this.writer = writer;
        }

        public void WriteHeader()
        {
            string line = "Duration (mm:ss),Scan Type,Ballot Cast Status,Simio Input (seconds)";
            line += this.fileName.Length > 0 ? ",File Name" : "";
            this.writer.WriteLineArr(line.Split(','));
        }

        private string[] GetElements(string line)
        {
            string[] elements = line.Split(',');
            for (int i = 0; i < elements.Length; i++)
            {
                elements[i] = elements[i].Trim();
            }
            return elements;
        }
        
        private DateTime GetTimestamp(string line)
        {
            string[] elements = GetElements(line);
            return DateTime.ParseExact(elements[1] + " " + elements[2], "MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private int GetCode(string line)
        {
            string[] elements = GetElements(line);
            return int.Parse(elements[0]);
        }

        private TimeSpan getDuration(string line2, string line1)
        {
            return GetTimestamp(line2) - GetTimestamp(line1);
        }

        private void WriteLineArr(string[] lineArr)
        {
            writer.WriteLineArr(this.fileName.Length > 0 ? Util.AppendToArray(lineArr, fileName) : lineArr, fieldTypes);
        }

        public void Done()
        {
            for (int i = 1; i < lines.Count; i++)
            {
                if ((i + 1) < lines.Count && GetCode(lines[i]) == votingStartedCode
                    && GetCode(lines[i + 1]) != blankBallotCode
                    && GetCode(lines[i + 1]) != overvotedBallotCode)
                {
                    string[] outputArr = new string[4];
                    outputArr[0] = getDuration(lines[i + 1], lines[i]).ToString(@"mm\:ss");
                    outputArr[3] = ((int)getDuration(lines[i + 1], lines[i]).TotalSeconds).ToString();
                    if (GetCode(lines[i + 1]) != votingCompleteCode)
                    {
                        outputArr[2] = "Unsuccessful";
                        outputArr[1] = GetElements(lines[i + 1])[6];
                    } else
                    {
                        outputArr[2] = "Successful";
                        outputArr[1] = "No Error";
                    }
                    i++;
                    WriteLineArr(outputArr);
                } else if ((i + 2) < lines.Count && GetCode(lines[i]) == votingStartedCode
                    && GetCode(lines[i + 2]) == votingCompleteCode)
                {
                    string[] outputArr = new string[4];
                    outputArr[0] = getDuration(lines[i + 2], lines[i]).ToString(@"mm\:ss");
                    outputArr[3] = ((int)getDuration(lines[i + 2], lines[i]).TotalSeconds).ToString();
                    outputArr[2] = "Successful";
                    outputArr[1] = GetElements(lines[i + 1])[6];
                    i += 2;
                    WriteLineArr(outputArr);
                } else if ((i + 1) < lines.Count && GetCode(lines[i]) == ballotJamCode
                    && GetCode(lines[i - 1]) != votingStartedCode && GetCode(lines[i + 1]) == jamClearedCode)
                {
                    string[] outputArr = new string[4];
                    outputArr[0] = getDuration(lines[i + 1], lines[i]).ToString(@"mm\:ss");
                    outputArr[3] = ((int)getDuration(lines[i + 1], lines[i]).TotalSeconds).ToString();
                    outputArr[2] = "Jam";
                    outputArr[1] = GetElements(lines[i])[6];
                    i++;
                    WriteLineArr(outputArr);
                } else if ((i + 1) < lines.Count && GetCode(lines[i]) == shutDownCode
                    && GetCode(lines[i - 1]) == unknownCode && GetCode(lines[i + 1]) == votingModeCode)
                {
                    string[] outputArr = new string[4];
                    outputArr[0] = getDuration(lines[i + 2], lines[i]).ToString(@"mm\:ss");
                    outputArr[3] = ((int)getDuration(lines[i + 2], lines[i]).TotalSeconds).ToString();
                    outputArr[2] = "Shutdown";
                    outputArr[1] = GetElements(lines[i])[6];
                    i++;
                    WriteLineArr(outputArr);
                }
            }
        }
    }
}
