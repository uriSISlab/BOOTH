using BOOTH.LogProcessors;
using BOOTH.LogProcessors.VSAP_BMD;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;

namespace BOOTH.LogProcessors.VSAP_BMD
{
    class VSAPBMD_Processor : ILogProcessor
    {
        private const string loadingBallotLog = "Loading Ballot";
        private const string languageSelectedLog = "Language Selected";
        private const string removedBallotLog = "Voter removed ballot before read by BMD";
        private const string ballotActivatedLog = "Ballot Activated and User session is ended";
        private const string printedBallotLog = "Printed ballot successfully";
        private const string castBallotLog = "Casted ballot successfullly"; // Typo in "sucessfully" as it appears in logs
        private const string removedPrintedBallotLog = "Ballot removed after printing";
        private const string provisionalBallotEjectedLog = "Provisonal Ballot ejected"; // Typo in "provisional" as it appears in logs
        private const string pollPassScannedLog = "poll-pass successfully scanned";
        private const string votingSessionLockedLog = "voting session locked after timeout done (Ballot not in BMD)";
        private const string errorScanningBPMLog = "Error scanning BPM - BPM not present";
        private const string quitVotingLog = "Returning ballot - quit voting";
        private const string startLog = "screen diagnostics Successful";
            
        // Enum to represent states of the VSAP BMD
        enum BMDState
        {
            // Initial state
            INIT,
            // Loading state is entered after the ballot loading has begun
            Loading,
            // Ballot has been activated, user can now vote (or cast their vote
            // if ballot is already voted in)
            Activated,
            // Ballot has been printed
            Printed,
            // An out-of-place removed ballot log has occured
            UnexpectedRemovedBallot
        }

        private string fileName;
        private DateTime startTime;
        private BMDState state;
        private bool pollPassUsed;
        private IOutputWriter writer;

        public VSAPBMD_Processor()
        {
            ClearState();
        }

        private void ClearState()
        {
            // TODO differentiate between state machine state and i/o state
            this.fileName = "";
            this.pollPassUsed = false;
            this.state = BMDState.INIT;
        }

        private void WriteLineToWriter(string line)
        {
            string[] lineArr = line.Split(new string[] { ", " }, StringSplitOptions.None);
            if (fileName.Length > 0)
            {
                lineArr = Util.AppendToArray(lineArr, fileName);
            }
            FieldType[] fieldTypes = null;
            if (lineArr[0] != "-")
            {
                fieldTypes = new FieldType[] { FieldType.TIMESPAN_MMSS }; 
            }
            writer.WriteLineArr(lineArr, fieldTypes);
        }

        private void WriteBallotRemovedRecordNoTime()
        {
            this.WriteLineToWriter("-, Voter removed ballot before read by BMD, Unsuccessful, -");
        }

        private void WriteBallotRemovedRecord(string duration)
        {
            this.WriteLineToWriter(duration + ", Voter removed ballot before read by BMD, Unsuccessful, -");
        }

        private void WriteBallotCastRecord(string duration, bool printed, bool pollPassUsed)
        {
            string outline = duration;
            if (printed)
            {
                outline += ", Ballot printed and cast, Successful";
                outline += pollPassUsed ? ", Yes" : ", No";
            } else
            {
                outline += ", Pre-printed ballot cast, Successful, -";
            }
            this.WriteLineToWriter(outline);
        }

        private void WriteProvisionalBallotEjectedRecord(string duration, bool printed, bool pollPassUsed)
        {
            string outline = duration;
            if (printed)
            {
                outline += ", Provisional ballot printed and ejected, Successful";
                outline += pollPassUsed ? ", Yes" : ", No";
            } else
            {
                outline += ", Pre-printed provisional ballot ejected, Successful, -";
            }
            this.WriteLineToWriter(outline);
        }

        private void WritePrintedBallotRemovedRecord(string duration, bool pollPassUsed)
        {
            string outline = duration + ", Ballot printed and removed, Unsuccessful";
            outline += pollPassUsed ? ", Yes" : ", No";
            this.WriteLineToWriter(outline);
        }

        private void WriteVotingTimedOutLog(string duration, bool pollPassUsed)
        {
            string outline = duration + ", Voting session timed out, Unsuccessful";
            outline += pollPassUsed ? ", Yes" : ", No";
            this.WriteLineToWriter(outline);
        }

        private void WriteBPMScanErrorLog(string duration)
        {
            this.WriteLineToWriter(duration + ", BPM Scan Error, Unsuccessful");
        }

        private void WriteQuitVotingLog(string duration, bool pollPassUsed)
        {
            string outline = duration + ", Voter quit voting, Unsuccessful";
            outline += pollPassUsed ? ", Yes" : ", No";
            this.WriteLineToWriter(outline);
        }

        private void WriteMachineRestartedLog(string duration, bool pollPassUsed)
        {
            string outline = duration + ", Voting machine restarted unexpectedly, Unsuccessful";
            outline += pollPassUsed ? ", Yes" : ", No";
            this.WriteLineToWriter(outline);
        }

        private DateTime ParseDate(string dateString)
        {
            return DateTime.ParseExact(dateString, "yyyy-MM-ddTHH:mm:ss.fffZ",
                System.Globalization.CultureInfo.InvariantCulture);
        }

        public string GetSeparator()
        {
            return "|";
        }

        public bool IsThisLog(Worksheet sheet)
        {
            return sheet.Range["D1"].Text.ToString().Trim() == "Logger.js-Loading page-Manual Diagnostic Status";
        }

        public void ReadLine(string line)
        {
            DateTime thisTime;
            string thisLog;
            string[] elements = line.Split('|');
            
            if (elements.Length == 7)
            {
                thisTime = this.ParseDate(elements[1]);
                thisLog = elements[6];
                if (this.state != BMDState.INIT && this.state != BMDState.UnexpectedRemovedBallot && startTime != null)
                {
                    if (Util.GetDifferenceMinutes(startTime, thisTime) > 60)
                    {
                        // A more than 60 minute difference probably indicates something suspicious.
                        // So we reset the state here.
                        this.state = BMDState.INIT;
                    }
                }

                switch (this.state)
                {
                    case BMDState.INIT:
                        switch (thisLog.Trim())
                        {
                            case loadingBallotLog:
                                this.startTime = thisTime;
                                this.state = BMDState.Loading;
                                this.pollPassUsed = false;
                                break;
                            case removedBallotLog:
                                this.state = BMDState.UnexpectedRemovedBallot;
                                break;
                        }
                        break;
                    case BMDState.UnexpectedRemovedBallot:
                        switch (thisLog.Trim())
                        {
                            case loadingBallotLog:
                                // Since this loading ballot log appears after an unexpected removed ballot log,
                                // we will assume that this log line is the one that should have come before
                                // the unexpected one we encountered before.
                                this.WriteBallotRemovedRecordNoTime();
                                this.state = BMDState.INIT;
                                break;
                        }
                        break;
                    case BMDState.Loading:
                        switch (thisLog.Trim())
                        {
                            case loadingBallotLog:
                                // I wfe encounter another "loading" log at this state,
                                // the first one most probably came after a mis-ordered one
                                this.WriteBallotRemovedRecordNoTime();
                                this.pollPassUsed = false;
                                this.startTime = thisTime;
                                break;
                            case removedBallotLog:
                                // This means the ballot was removed from the machine before it could
                                // be read and activated. We need to record it and reset state.
                                this.WriteBallotRemovedRecord(Util.GetTimeDifference(startTime, thisTime));
                                this.state = BMDState.INIT;
                                break;
                            case ballotActivatedLog:
                                this.state = BMDState.Activated;
                                break;
                            case errorScanningBPMLog:
                                this.WriteBPMScanErrorLog(Util.GetTimeDifference(startTime, thisTime));
                                this.state = BMDState.INIT;
                                break;
                            case startLog:
                                this.WriteMachineRestartedLog(Util.GetTimeDifference(startTime, thisTime), false);
                                this.state = BMDState.INIT;
                                break;
                            case quitVotingLog:
                                this.WriteQuitVotingLog(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                        }
                        break;
                    case BMDState.Activated:
                        switch (thisLog.Trim())
                        {
                            case pollPassScannedLog:
                                this.pollPassUsed = true;
                                break;
                            case printedBallotLog:
                                this.state = BMDState.Printed;
                                break;
                            case castBallotLog:
                                // If the ballot was cast withut being printed in this transaction.
                                // This means a pre-printed ballot was inserted.
                                this.WriteBallotCastRecord(Util.GetTimeDifference(startTime, thisTime), false, this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case provisionalBallotEjectedLog:
                                this.WriteProvisionalBallotEjectedRecord(Util.GetTimeDifference(startTime, thisTime), false, this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case votingSessionLockedLog:
                                this.WriteVotingTimedOutLog(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case quitVotingLog:
                                this.WriteQuitVotingLog(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case startLog:
                                this.WriteMachineRestartedLog(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case languageSelectedLog:
                                this.WriteQuitVotingLog(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                        }
                        break;
                    case BMDState.Printed:
                        switch (thisLog.Trim())
                        {
                            case removedPrintedBallotLog:
                                this.WritePrintedBallotRemovedRecord(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case castBallotLog:
                                this.WriteBallotCastRecord(Util.GetTimeDifference(startTime, thisTime), true, this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case provisionalBallotEjectedLog:
                                this.WriteProvisionalBallotEjectedRecord(Util.GetTimeDifference(startTime, thisTime), true, this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            case startLog:
                                this.WriteMachineRestartedLog(Util.GetTimeDifference(startTime, thisTime), this.pollPassUsed);
                                this.state = BMDState.INIT;
                                break;
                            // TODO Find out whether provisional ballots can be cast just after printing
                        }
                        break;
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
            if (fileName.Length > 0)
            {
                writer.WriteLine("Duration (mm:ss)", "Scan Type", "Ballot Cast Status", "Poll Pass Used", "Filename");
            } else
            {
                writer.WriteLine("Duration (mm:ss)", "Scan Type", "Ballot Cast Status", "Poll Pass Used");
            }
        }

        public void Done()
        {
        }

        public string GetUniqueTag()
        {
            return VSAPBMD_Summarizer.MACHINE_TYPE_TAG;
        }
    }
}
