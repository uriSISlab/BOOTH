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

namespace BOOTH.LogProcessors.Dominion_ICE
{
    public class DICE_Processor : ILogProcessor
    {
        private const string s_ballotInserted = "[Voting] A ballot has been inserted into the unit.";
        private const string s_invokingBallotReview = "[Voting] Invoking ballot review for the current ballot.";
        private const string s_ballotCast = "[Voting] Ballot successfully cast and dropped into ballot box";
        private const string s_endOfSession = "[Voting] The end of the current session. The system is ready to accept a new ballot.";
        private const string s_accessibleStarted = "[Admin] Accessible Voting (Contest by Contest) started";
        private const string s_accessibleBallotInserted = "[Accessible Voting] A ballot has been inserted into the unit";
        private const string s_accessibleBallotAccepted = "[Accessible Voting] A ballot has been accepted by the system.";
        private const string s_accessibleMarkingCompleted = "[Accessible Voting] A ballot marking has been completed.";
        private const string s_ballotMisread = "[Pixel Count] Ballot misread.";
        private const string s_paperJam = "[Scanner] Error scanning ballot: Possible paper jam. Code:";
        private const string s_noManifestation = "[Pixel Count] No ballot manifestation for determined ballot Id";

        private enum DICEState {
            // The string device is ready to accept a new ballot
            Ready,
            // A ballot has just been inserted into the device for scanning
            BallotInserted,
            // A ballot was just cast into the machine
            BallotCast,
            // An accessible voting session has been started
            AccessibleStarted,
            // A ballot has been inserted for accessible voting
            AccessibleBallotInserted,
            // A ballot has been accepted for accessible voting
            AccessibleBallotAccepted
        }

        private DICEState state;
        private DateTime startTime;
        private IOutputWriter writer;
        private int misreads;
        private bool reviewed;
        private bool paperJam;
        private bool ballotNotRecognized;
        private DateTime lastTimestamp;
        private string fileName;

        public DICE_Processor()
        {
            ClearState();
        }

        private void ClearState()
        {
            this.state = DICEState.Ready;
            this.fileName = "";
            this.misreads = 0;
            this.reviewed = false;
            this.paperJam = false;
            this.ballotNotRecognized = false;
        }

        private void WriteToWriter(string[] lineArr)
        {
            // Write the given array to writer after possibly appending the filename
            FieldType[] fieldTypes = new FieldType[] { FieldType.TIMESPAN_MMSS, FieldType.DATETIME, FieldType.STRING,
                FieldType.STRING, FieldType.STRING, FieldType.STRING };
            writer.WriteLineArr(this.fileName.Length > 0 ? Util.AppendToArray(lineArr, this.fileName) : lineArr, fieldTypes);
        }

        private DateTime GetTimestampFromDICELine(string line)
        {
            return DateTime.ParseExact(line.Substring(0, 20), "dd MMM yyyy HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private void WriteBallotCastNormalRecord(DateTime startTimestamp, DateTime endTimestamp)
        {
            this.WriteToWriter(new string[] { Util.GetTimeDifference(startTimestamp, endTimestamp), endTimestamp.ToString(),
                "Ballot cast normally", this.misreads.ToString(), this.reviewed ? "Yes" : "No" });
        }

        private void WriteBallotNotCastRecord(DateTime startTimestamp, DateTime endTimestamp)
        {
            this.WriteToWriter(new string[] { Util.GetTimeDifference(startTimestamp, endTimestamp), endTimestamp.ToString(),
                "Ballot not cast", this.misreads.ToString(), this.reviewed ? "Yes" : "No" });
        }

        private void WritePaperJamRecord(DateTime startTimestamp, DateTime endTimestamp)
        {
            this.WriteToWriter(new string[] { Util.GetTimeDifference(startTimestamp, endTimestamp), endTimestamp.ToString(),
                "Paper jam when accepting ballot", this.misreads.ToString(), "-"});
        }

        private void WriteBallotNotRecognizedRecord(DateTime startTimestamp, DateTime endTimestamp)
        {
            this.WriteToWriter(new string[] { Util.GetTimeDifference(startTimestamp, endTimestamp), endTimestamp.ToString(),
                "Ballot not recognized", this.misreads.ToString(), "-"});
        }

        private void WriteAccessibleBallotMarkedRecord(DateTime startTimestamp, DateTime endTimestamp)
        {
            this.WriteToWriter(new string[] { Util.GetTimeDifference(startTimestamp, endTimestamp), endTimestamp.ToString(),
                "Accessible voting : ballot marked", "-" , "-"}); // TODO find out if accessible marking mode review is possible
        }

        private void WriteRecordAccordingToState(DateTime startTimestamp, DateTime endTimestamp)
        {
            switch (this.state)
            {
                case DICEState.BallotInserted:
                    if (this.paperJam)
                    {
                        this.WritePaperJamRecord(startTimestamp, endTimestamp);
                    } else if (this.ballotNotRecognized)
                    {
                        this.WriteBallotNotRecognizedRecord(startTimestamp, endTimestamp);
                    } else
                    {
                        this.WriteBallotNotCastRecord(startTimestamp, endTimestamp);
                    }
                    break;
                case DICEState.BallotCast:
                    this.WriteBallotCastNormalRecord(startTimestamp, endTimestamp);
                    break;
            }
        }

        public string GetSeparator()
        {
            return " ";
        }

        public bool IsThisLog(Worksheet sheet)
        {
            return sheet.Cells[1, 2].Text.ToString().Contains("Logging service initialized");
        }

        public void ReadLine(string line)
        {
            // The timestamp is 20 characters long
            if (line.Length < 21)
            {
                // There is nothing useful in the line, do nothing
                return;
            }
            // Get the position of the first colon after the timestamp
            int colPos = 21 + line.Substring(21).IndexOf(":");
            if (colPos == 20)
            {
                // Do nothing if the colon was not found
                return;
            }

            DateTime thisTime = GetTimestampFromDICELine(line);
            string rest = line.Substring(colPos + 1).Trim();

            switch (rest)
            {
                case s_accessibleStarted:
                    this.state = DICEState.AccessibleStarted;
                    this.startTime = thisTime;
                    break;
                case s_accessibleMarkingCompleted:
                    this.state = DICEState.Ready;
                    this.WriteAccessibleBallotMarkedRecord(startTime, thisTime);
                    break;
                case s_ballotInserted:
                    if (Util.GetDifferenceMinutes(this.lastTimestamp, thisTime) > 2)
                    {
                        // If there's a large difference in time here, the earlier session
                        // was probably abandoned due to something. So we clear the state.
                        this.ClearState();
                    }
                    this.state = DICEState.BallotInserted;
                    this.startTime = thisTime;
                    break;
                case s_ballotCast:
                    this.state = DICEState.BallotCast;
                    break;
                case s_endOfSession:
                    this.WriteRecordAccordingToState(this.startTime, thisTime);
                    this.ClearState();
                    break;
                case s_ballotMisread:
                    this.misreads++;
                    break;
                case s_invokingBallotReview:
                    this.reviewed = true;
                    break;
                case s_noManifestation:
                    this.ballotNotRecognized = true;
                    break;
                default:
                    if (rest.Contains(s_paperJam))
                    {
                        this.paperJam = true;
                    }
                    break;
            }

            this.lastTimestamp = thisTime;
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
            // Write the header
            string line = "Duration,Timestamp,Event,Misreads,Ballot Reviewed";
            line += fileName.Length > 0 ? ",Filename" : "";
            writer.WriteLineArr(line.Split(','));
        }

        public void Done()
        {
        }

        public string GetUniqueTag()
        {
            return "DICE";
        }
    }
}
