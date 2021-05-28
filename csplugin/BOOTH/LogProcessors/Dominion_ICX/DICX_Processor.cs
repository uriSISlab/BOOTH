using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;

namespace BOOTH
{
    class DICX_Processor : ILogProcessor
    {
        
        private enum DICXState
        {
            // The device is ready to accept a new ballot
            Ready,
            // A new voting session has been started
            Started,
            // Voter has voted and is reviewing ballot
            BallotReview,
            // Ballot is being cast
            BallotCasting
        }

        private const string s_startingVotingSession = "Starting new voting session.";
        private const string s_ballotPresented = "Ballot is presented to voter.";
        private const string s_ballotReview = "Ballot review";
        private const string s_prepareBallot = "Prepare ballot for cast.";
        private const string s_ballotCast = "Ballot cast successfully";

        private DICXState state;
        private DateTime startTime;
        private IOutputWriter writer;
        private String fileName;

        public DICX_Processor()
        {
            this.ClearState();
        }

        private void ClearState()
        {
            this.state = DICXState.Ready;
            this.fileName = "";
        }
        
        private void WriteToWriter(string[] lineArr)
        {
            FieldType[] fieldTypes = new FieldType[] { FieldType.TIMESPAN_MMSS,
                FieldType.DATETIME, FieldType.STRING, FieldType.STRING };
            // Write the given array to writer after possibly appending the filename
            writer.WriteLineArr(this.fileName.Length > 0 ? Util.AppendToArray(lineArr, fileName) : lineArr,
                fieldTypes);
        }

        private void WriteBallotCastNormalRecord(DateTime startTimestamp, DateTime endTimestamp)
        {
            this.WriteToWriter(new string[] {Util.GetTimeDifference(startTimestamp, endTimestamp),
                endTimestamp.ToString(), "Ballot cast normally"});
        }

        public string GetSeparator()
        {
            return " - ";
        }

        public bool IsThisLog(Worksheet sheet)
        {
            Range idRange = sheet.UsedRange.Find(What: "Audit Log file is saved.",
                LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlWhole, SearchOrder: XlSearchOrder.xlByRows,
                SearchDirection: XlSearchDirection.xlNext, MatchCase: false);
            return idRange != null;
        }

        public void ReadLine(string line)
        {
            if (line.Length < 24)
            {
                // There is nothing useful in the line, do nothing
                return;
            }

            // Check the position of the timestamp-log divider
            if (!line.Substring(1).Contains(" - "))
            {
                // Do nothing if the divider was not found
                return;
            }

            // Timestamp is in the first 19 characters
            DateTime thisTime = DateTime.ParseExact(line.Substring(0, 19), "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            string rest = line.Substring(22).Trim(); // Next three characters are " - "

            switch (rest)
            {
                case s_startingVotingSession:
                    this.state = DICXState.Started;
                    this.startTime = thisTime;
                    break;
                case s_ballotReview:
                    this.state = DICXState.BallotReview;
                    break;
                case s_prepareBallot:
                    this.state = DICXState.BallotCasting;
                    break;
                case s_ballotCast:
                    this.state = DICXState.Ready;
                    this.WriteBallotCastNormalRecord(this.startTime, thisTime);
                    break;
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
            writer.WriteLineArr(("Duration,Timestamp,Event" + (fileName.Length > 0 ? ",Filename" : "")).Split(','));
        }

        public void Done()
        {
        }
    }
}
