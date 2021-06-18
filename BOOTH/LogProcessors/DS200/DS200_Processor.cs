using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace BOOTH.LogProcessors.DS200
{
    class DS200_Processor : ILogProcessor
    {

        private const int votingStartedCode = 1004115;
        private const int blankBallotCode = 1004113;
        private const int overvotedBallotAcceptedCode = 1004111;
        private const int votingCompleteCode = 1004022;
        private const int ballotJamCode = 3013004;
        private const int jamClearedCode = 1004328;
        private const int shutDownCode = 1004016;
        private const int lidClosedCode = 1004163; // ??
        private const int votingModeCode = 1004056;

        private const string s_votingStarted = "Vote Session Started";
        private const string s_blankBallotAccepted = "Voter Accepted Blank Ballot";
        private const string s_overVotedBallotAccepted = "Voter Accepted Overvoted Ballot";
        private const string s_blankBallotRejected = "Voter Rejected Blank Ballot";
        private const string s_overVotedBallotRejected = "Voter Rejected Overvoted Ballot";
        private const string s_votingComplete = "Voting session complete";
        private const string s_ballotJamCheckPath = "Ballot Jam.  Please check the paper path.";
        private const string s_ballotJamCleared = "Ballot jam cleared";
        private const string s_shutdown = "Shutdown initiated";
        private const string s_lidClosed = "Lid Closed. Waiting for automatic shutdown.";
        private const string s_votingMode = "Entering voting mode";
        private const string s_systemError = "System Error - Contact Election Official.";
        private const string s_exitingAdminMenus = "Exiting Administration Menus";
        private const string s_multipleBallots = "Multiple ballots were detected. "
            + "Please remove ballots and insert them one ballot at a time. "
            + "Ensure your ballot is not folded or damaged.";
        private const string s_reinsertOpposite = "Ballot Could Not Be Read. "
            + "Please remove your ballot and re-insert the opposite end first.";
        private const string s_ballotRemoved = "Ballot was removed during scanning. "
            + "Please re-insert the ballot completely.";
        private const string s_removeStubs = "Error scanning ballot. "
            + "Please remove your ballot and re-insert the opposite end first. "
            + "Ensure all stubs are removed from the ballot.";
        private const string s_notInsertedFarEnough = "Ballot was not inserted far enough. "
            + "Please remove your ballot and re-insert it completely.";
        private const string s_machineNotProgrammed = "Voting Machine Not Programmed For Your Ballot";
        private const string s_ballotJamReinsert = "Ballot Jam. Please remove ballot and re-insert.";
        private const string s_ballotJamCheckPath2 = "Ballot Jam. Please check the paper path.";
        private const string s_ballotTooShort = "Ballot too short Please remove ballot.";
        private const string s_autoRejectBallot = "Automatically rejected Ballot with Unreadable mark";
        private const string s_unprocessedElement = "Unprocessed ballot element. Ballot cannot be scanned.";

        private enum State
        {
            Ready,
            VotingStarted,
            BallotJammed,
        }

        private readonly int[] recognizedCodes = new int[]
        {
            1004115, 1004163, 1004016, 1004056,
            1004022, 1004111, 1004113, 3013004,
            1004328,

            3013006,    // System Error - Contact Election Official
            1004138,    // Exiting Administration Menus
            3013005,    // Multiple Ballots Detected
            3003337,    // Ballot could not be read. Reinsert opposite end.
            3013001,    // Ballot removed during scanning. Reinsert ballot.
            3013008,    // Error scanning ballot. Reinsert opposite end, remove all stubs.
            3013002,    // Ballot not inserted far enough. Reinsert.
            7003009,    // Voting machine not programmed for your ballot.
            3013003,    // Ballot jam. Reinsert.
            3013007,    // Ballot jam. Check paper path.
            3013009,    // Ballot too short. Remove ballot.
            3003335,    // Ballot could not be read. Reinsert opposite end.
            3003336,    // Ballot could not be read. Reinsert opposite end.
            3003339,    // Ballot could not be read. Reinsert opposite end.
            3003340,    // Ballot could not be read. Reinsert opposite end.
            3003318,    // Ballot could not be read. Reinsert opposite end.
            3003341,    // Ballot could not be read. Reinsert opposite end.
            1004122,    // Rejected ballot with unreadable mark.
            1004112,    // Voter rejected overvoted ballot
            1004114,    // Voter rejected blank ballot
        };

        private readonly FieldType[] fieldTypes = new FieldType[] { FieldType.TIMESPAN_MMSS, FieldType.STRING,
            FieldType.STRING, FieldType.INTEGER};

        private string fileName;
        private IOutputWriter writer;
        private readonly List<string> lines;
        private State state;
        private DateTime startTimestamp;

        public DS200_Processor()
        {
            this.fileName = "";
            this.lines = new List<string>();
            this.state = State.Ready;
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
            string[] elements = this.GetElements(line);
            DateTime timestamp = DateTime.ParseExact(elements[1] + " " + elements[2], "MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            if (this.state == State.Ready)
            {
                if (line.Contains(s_votingStarted))
                {
                    this.startTimestamp = timestamp;
                    this.state = State.VotingStarted;
                }
            } else if (this.state == State.VotingStarted)
            {
                if (line.Contains(s_votingComplete))
                {
                    this.WriteSuccessfulRecord(timestamp);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_systemError))
                {
                    this.WriteRecord(timestamp, "System error", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_blankBallotRejected))
                {
                    this.WriteRecord(timestamp, "Blank ballot rejected", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_overVotedBallotRejected))
                {
                    this.WriteRecord(timestamp, "Overvoted ballot rejected", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_ballotJamReinsert) ||
                  line.Contains(s_ballotJamCheckPath) || line.Contains(s_ballotJamCheckPath2))
                {
                    this.WriteRecord(timestamp, "Ballot jam", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_multipleBallots))
                {
                    this.WriteRecord(timestamp, "Multiple ballots detected", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_reinsertOpposite))
                {
                    this.WriteRecord(timestamp, "Ballot could not be read", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_removeStubs))
                {
                    this.WriteRecord(timestamp, "Error scanning ballot", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_ballotRemoved))
                {
                    this.WriteRecord(timestamp, "Ballot removed during scan", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_notInsertedFarEnough))
                {
                    this.WriteRecord(timestamp, "Ballot not inserted far enough", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_machineNotProgrammed))
                {
                    this.WriteRecord(timestamp, "Voting machine not programmed for ballot", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_ballotTooShort))
                {
                    this.WriteRecord(timestamp, "Ballot too short", false);
                    this.state = State.Ready;
                }
                else if (line.Contains(s_shutdown))
                {
                    this.WriteRecord(timestamp, "Machine shutdown", false);
                    this.state = State.Ready;
                } else if (line.Contains(s_autoRejectBallot))
                {
                    this.WriteRecord(timestamp, "Rejected ballot with unreadable mark", false);
                    this.state = State.Ready;
                } else if (line.Contains(s_unprocessedElement))
                {
                    this.WriteRecord(timestamp, "Unprocessed ballot element", false);
                    this.state = State.Ready;
                } else if (line.Contains(s_votingStarted))
                {
                    System.Diagnostics.Debug.WriteLine("Unexpected end to voting session!");
                    if (this.fileName != null)
                    {
                        System.Diagnostics.Debug.WriteLine("In file " + this.fileName);
                    }
                    System.Diagnostics.Debug.WriteLine(line);
                    this.startTimestamp = timestamp;
                    this.state = State.VotingStarted;
                }
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

        private void WriteSuccessfulRecord(DateTime currentTimestamp)
        {
            this.WriteRecord(currentTimestamp, "Voting session complete", true);
        }

        private void WriteRecord(DateTime currentTimestamp, string eventStr, bool successful)
        {
            string[] outputArr = new string[4];
            TimeSpan delta = (currentTimestamp - this.startTimestamp);
            outputArr[0] = delta.ToString(@"mm\:ss");
            outputArr[1] = eventStr;
            outputArr[2] = successful ? "Successful" : "Unsuccessful";
            outputArr[3] = ((int)delta.TotalSeconds).ToString();
            this.WriteLineArr(outputArr);
        }

        private string[] GetElements(string line)
        {
            // TODO: This will not work if there are escaped commas inside a field
            string[] elements = line.Split(',');
            for (int i = 0; i < elements.Length; i++)
            {
                elements[i] = elements[i].Trim();
            }
            return elements;
        }

        private void WriteLineArr(string[] lineArr)
        {
            writer.WriteLineArr(this.fileName.Length > 0 ? Util.AppendToArray(lineArr, fileName) : lineArr, fieldTypes);
        }

        public void Done()
        {
            // this.writer.Flush();
        }

        public string GetUniqueTag()
        {
            return DS200_Summarizer.MACHINE_TYPE_TAG;
        }
    }
}
