using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BOOTH.LogProcessors.PollPad
{
    class PollPad_Processor : ILogProcessor
    {
        private static readonly string s_selectCheckinIdMethod = "SELECT CHECKIN IDENTIFICATION METHOD";
        private static readonly string s_manual = "MANUAL";
        private static readonly string s_scanId = "SCAN ID";
        private static readonly string s_lookupTableSelection = "VOTER LOOKUP TABLE SELECTION";
        private static readonly string s_cancelVbm = "ISSUE BALLOT - CANCEL VBM";
        private static readonly string s_assistanceRequired = "ADDED AND SIGNED: ASSISTANCE REQUIRED";
        private static readonly string s_registrationTapped = "REGISTRATION BUTTON TAPPED";
        private static readonly string s_addedVoter = "ADDED VOTER";
        private static readonly string s_editedVoter = "EDITED VOTER";
        private static readonly string s_checkinVoter = "CHECK IN VOTER";
        private static readonly string s_spoiledBallot = "SPOILED BALLOT";
        private static readonly string s_cancelCheckin = "CANCEL CHECKIN | Cancelling checkin";
        private static readonly string s_voterSearch = "VOTER LOOKUP SEARCH FOR VOTER | TIME TAKEN IN SECONDS:";
        private static readonly string s_toIdSelectionScreen = "TO: ID SELECTION FOR LOOKUP SCREEN";
        private static readonly FieldType[] recordFieldTypes =
            {
                FieldType.TIMESPAN_MMSS, FieldType.STRING, FieldType.DATETIME, FieldType.STRING,
                FieldType.STRING, FieldType.INTEGER, FieldType.STRING, FieldType.STRING
            };

        enum PollPadState
        {
            INIT,
            STARTED,
            SELECTED,
            REGISTERING,
            JUST_ADDED,
            JUST_EDITED
        }

        enum RecordType
        {
            // Voter was checked in normally
            CHECKED_IN_NORMAL,
            // Voter was checked in after adding their data to the records
            CHECKED_IN_AFTER_ADD,
            // Voter was checked in after adding their data was edited
            CHECKED_IN_AFTER_EDIT,
            // Check-in process was abandoned after voter lookup(s)
            ABANDONED_AFTER_SEARCH,
            // Check-in process was abandoned after voter was selected from the list
            ABANDONED_AFTER_SELECTING,
            // Process of registering (adding) the voter was abandoned before completion
            REGISTRATION_INCOMPLETE,
            // Check-in process was abandoned after adding voter.
            ABANDONED_AFTER_ADDING,
            // Check-in process was abandoned after editing voter data.
            ABANDONED_AFTER_EDITING,
            // Voter's ballot was spoiled
            BALLOT_SPOILED,
            // Check-in was canceled
            CANCEL_CHECK_IN
        }

        private IOutputWriter writer;
        private string fileName;

        private PollPadState state = PollPadState.INIT;
        private bool scanIdLookup = false;
        private DateTime startTime;
        private DateTime endTime;
        private bool vbmCancelled = false;
        private bool assistanceRequired = false;
        private uint searches = 0;
        private bool durationHighConfidence = true;
        private DateTime lastTimestamp;

        private string[] recordLine = new string[8];

        private static string GetPrettyStringForRecordType(RecordType recordType)
        {
            switch (recordType)
            {
                case RecordType.CHECKED_IN_NORMAL:
                    return "Normal check-in";
                case RecordType.CHECKED_IN_AFTER_ADD:
                    return "Checked in after voter added";
                case RecordType.CHECKED_IN_AFTER_EDIT:
                    return "Checked in after voter edited";
                case RecordType.ABANDONED_AFTER_SEARCH:
                    return "Check-in abandoned after voter search";
                case RecordType.ABANDONED_AFTER_SELECTING:
                    return "Check-in abandoned after voter selected";
                case RecordType.ABANDONED_AFTER_ADDING:
                    return "Check-in abandoned after voter added";
                case RecordType.ABANDONED_AFTER_EDITING:
                    return "Check-in abandoned after voter edited";
                case RecordType.REGISTRATION_INCOMPLETE:
                    return "Voter registration did not complete";
                case RecordType.BALLOT_SPOILED:
                    return "Voter's ballot was spoiled";
                case RecordType.CANCEL_CHECK_IN:
                    return "Check-in cancelled";
                default:
                    return "Unknown";
            }
        }

        public string GetSeparator()
        {
            return " | ";
        }

        public string GetUniqueTag()
        {
            return PollPad_Summarizer.MACHINE_TYPE_TAG;
        }

        public bool IsThisLog(Worksheet sheet)
        {
            DynamicSheetReader reader = new DynamicSheetReader(sheet, this.GetSeparator());
            while (!reader.NoMoreLines())
            {
                string line = reader.ReadLine();
                if (PollPad_Importer.timestampRegex.IsMatch(line) && line.Contains("NAVIGATION"))
                {
                    return true;
                }
            }
            return false;
        }

        public void ReadLine(string line)
        {
            MatchCollection matches = PollPad_Importer.timestampRegex.Matches(line);
            if (matches.Count == 0)
            {
                // System.Diagnostics.Debug.WriteLine("Timestamp-less line passed to PollPad_Processor!");
                return;
            }
            // Extract and parse timestamp
            DateTime timestamp = DateTime.ParseExact(matches[0].Value, @"MMM d, yyyy \a\t h:mm:ss tt", CultureInfo.InvariantCulture);
            // Reset the state if there's a large (> 15 minute) gap between the previous and the current timestamp
            if (this.lastTimestamp != Util.nullTimestamp && (timestamp - this.lastTimestamp).TotalMinutes > 15)
            {
                // Add an "abandoned" record if there is something to add
                this.AddAbandonedRecord(this.lastTimestamp);
                this.ClearState();
            }

            if (this.state == PollPadState.INIT)
            {
                if (line.Contains(s_selectCheckinIdMethod))
                {
                    // This is the normal start checkin flow. Set the state and other state variables.
                    this.SetStartingState(line, timestamp);
                } else if (line.Contains(s_registrationTapped))
                {
                    // If registration button was tapped while IDLE, the poll worker is trying to register
                    // a voter without looking them up first. Set state appropriately.
                    this.SetStartingState(line, timestamp, PollPadState.REGISTERING);
                } else if (line.Contains(s_voterSearch))
                {
                    // Read unexpected search log (before select checkin method log). The inferred duration
                    // might not be accurate.
                    this.SetStartingState(line, timestamp);
                    this.searches += 1;
                    this.durationHighConfidence = false;
                }
            } else if (this.state == PollPadState.STARTED)
            {
                if (line.Contains(s_selectCheckinIdMethod))
                {
                    // Previous checkin process ended without anything meaningful happening, so reset state.
                    this.SetStartingState(line, timestamp);
                } else if (line.Contains(s_toIdSelectionScreen))
                {
                    // The user went back, so reset state.
                    if (this.searches > 0)
                    {
                        // If some searches were done after the start, it's better to record an event.
                        this.AddAbandonedRecord(timestamp);
                    }
                    this.ClearState();
                } else if (line.Contains(s_lookupTableSelection))
                {
                    // We set an 'end time' value at this moment so that we get a better value for event duration
                    // in case the poll worker cancels the check-in process. (Apparently the log doesn't explicitly
                    // indicate such a cancellation and we have to infer it from the next check-in being started).
                    this.endTime = timestamp;
                    this.state = PollPadState.SELECTED;
                } else if (line.Contains(s_checkinVoter))
                {
                    // Check-in completes right after it has started when checking in with an ID (as opposed to
                    // manual lookup).
                    this.AddRecord(RecordType.CHECKED_IN_NORMAL, timestamp, null, this.vbmCancelled, this.assistanceRequired);
                    this.ClearState();
                } else if (line.Contains(s_spoiledBallot))
                {
                    // The voter's ballot has been spoiled.
                    this.AddRecord(RecordType.BALLOT_SPOILED, timestamp);
                    this.ClearState();
                } else if (line.Contains(s_cancelCheckin))
                {
                    // A previous check-in was cancelled.
                    this.AddRecord(RecordType.CANCEL_CHECK_IN, timestamp);
                    this.ClearState();
                } else if (line.Contains(s_assistanceRequired))
                {
                    // THe presence of this string indicates that the voter required assistance.
                    this.assistanceRequired = true;
                } else if (line.Contains(s_registrationTapped))
                {
                    // Registration (voter add) process was started after the start of check-in process.
                    this.state = PollPadState.REGISTERING;
                } else if (line.Contains(s_voterSearch))
                {
                    // Search query for voter detected, increment count.
                    this.searches += 1;
                    if ((timestamp - this.startTime).TotalMinutes > 3)
                    {
                        // If there's more than 3 minutes of difference between 'start' and search, assume
                        // that the "start" happened before the voter arrived and set a new start time.
                        this.durationHighConfidence = false;
                        this.startTime = timestamp;
                    }
                }
            } else if (state == PollPadState.SELECTED)
            {
                if (line.Contains(s_selectCheckinIdMethod))
                {
                    // Previous check-in process was abandoned and a new one started. Write a record of the
                    // previous one. Again, the duration here will be unreliable.
                    this.AddAbandonedRecord(this.endTime);
                    this.SetStartingState(line, timestamp);
                } else if (line.Contains(s_toIdSelectionScreen)) {
                    // The user went back, so reset state after writing a record.
                    this.AddAbandonedRecord(this.endTime);
                    this.ClearState();
                } else if (line.Contains(s_checkinVoter))
                {
                    // Voter checked in normally
                    this.AddRecord(RecordType.CHECKED_IN_NORMAL, timestamp, null, this.vbmCancelled,
                        this.assistanceRequired);
                    this.ClearState();
                } else if (line.Contains(s_spoiledBallot))
                {
                    // The voter's ballot has been spoiled.
                    this.AddRecord(RecordType.BALLOT_SPOILED, timestamp);
                    this.ClearState();
                } else if (line.Contains(s_cancelCheckin))
                {
                    // A previous check-in was cancelled.
                    this.AddRecord(RecordType.CANCEL_CHECK_IN, timestamp);
                    this.ClearState();
                } else if (line.Contains(s_voterSearch))
                {
                    // A re-search can happen even after selecting, so detect that.
                    this.searches += 1;
                } else if (line.Contains(s_cancelVbm))
                {
                    // Presence of this string indicates that the voter was a VBM voter and their mail
                    // ballot was cancelled.
                    this.vbmCancelled = true;
                } else if (line.Contains(s_assistanceRequired))
                {
                    // The presence of this string indicates the voter required assistance.
                    this.assistanceRequired = true;
                } else if (line.Contains(s_registrationTapped))
                {
                    // Registration (voter add) process was started after a voter was selected from search
                    // results.
                    this.state = PollPadState.REGISTERING;
                } else if (line.Contains(s_addedVoter))
                {
                    this.state = PollPadState.JUST_ADDED;
                } else if (line.Contains(s_editedVoter))
                {
                    this.state = PollPadState.JUST_EDITED;
                }
            } else if (this.state == PollPadState.REGISTERING)
            {
                if (line.Contains(s_addedVoter))
                {
                    // Voter data was added to the database.
                    this.state = PollPadState.JUST_ADDED;
                } else if (line.Contains(s_editedVoter))
                {
                    // Voter's data was edited in the database.
                    this.state = PollPadState.JUST_EDITED;
                } else if (line.Contains(s_selectCheckinIdMethod))
                {
                    // The registration process was abandoned and a new check-in process is starting.
                    // Add a record of the abandoned process. Note that the duration may be inaccurate
                    // in this case.
                    this.durationHighConfidence = false;
                    this.AddRecord(RecordType.REGISTRATION_INCOMPLETE, timestamp);
                    this.SetStartingState(line, timestamp);
                } else if (line.Contains(s_toIdSelectionScreen))
                {
                    // The user went back, so reset state after writing a record.
                    this.durationHighConfidence = false;
                    this.AddRecord(RecordType.REGISTRATION_INCOMPLETE, timestamp);
                    this.ClearState();
                }
            } else if (this.state == PollPadState.JUST_ADDED || this.state == PollPadState.JUST_EDITED)
            {
                if (line.Contains(s_checkinVoter))
                {
                    // A check-in after adding the voter's data.
                    RecordType recordType = this.state == PollPadState.JUST_ADDED ?
                        RecordType.CHECKED_IN_AFTER_ADD : RecordType.CHECKED_IN_AFTER_EDIT;
                    this.AddRecord(recordType, timestamp, null, this.assistanceRequired);
                    this.ClearState();
                } else if (line.Contains(s_assistanceRequired))
                {
                    // The 'assistance required' line may appear at this stage too so we check for it.
                    this.assistanceRequired = true;
                } else if (line.Contains(s_selectCheckinIdMethod))
                {
                    // A start of checkin lin here means the add/edit process was abandoned. Write record
                    // and reset state. Again, the duration may very well be inaccurate.
                    this.AddAbandonedRecord(timestamp);
                    this.SetStartingState(line, timestamp);
                } else if (line.Contains(s_toIdSelectionScreen))
                {
                    // The user went back, so reset state after writing a record.
                    this.AddAbandonedRecord(timestamp);
                    this.ClearState();
                } else if (line.Contains(s_voterSearch)) {
                    this.searches += 1;
                }
            }
            this.lastTimestamp = timestamp;
        }

        private void SetStartingState(string line, DateTime timestamp, PollPadState state=PollPadState.STARTED)
        {
            this.ClearState();
            this.state = state;
            this.startTime = timestamp;
            if (line.Contains(s_selectCheckinIdMethod))
            {
                this.scanIdLookup = line.Contains(s_scanId);
            }
        }

        private bool AddAbandonedRecord(DateTime endTime)
        {
            this.durationHighConfidence = false;
            if (this.state == PollPadState.STARTED && this.searches > 0)
            {
                this.AddRecord(RecordType.ABANDONED_AFTER_SEARCH, endTime);
            } else if (this.state == PollPadState.SELECTED)
            {
                this.AddRecord(RecordType.ABANDONED_AFTER_SELECTING, endTime, this.startTime, this.vbmCancelled,
                    this.assistanceRequired);
            } else if (this.state == PollPadState.JUST_ADDED)
            {
                this.AddRecord(RecordType.ABANDONED_AFTER_ADDING, endTime, null, this.vbmCancelled,
                    this.assistanceRequired);
            } else if (this.state == PollPadState.JUST_EDITED)
            {
                this.AddRecord(RecordType.ABANDONED_AFTER_EDITING, endTime, null, this.vbmCancelled,
                    this.assistanceRequired);
            } else
            {
                // Return false if a record wasn't added, otherwise true
                return false;
            }
            return true;
        }

        private void AddRecord(RecordType recordType, DateTime endTime, DateTime? startTime=null, bool vbmCancelled=false,
            bool assistanceRequired=false)
        {
            if (startTime == null) {
                startTime = this.startTime;
            }
            TimeSpan delta = (TimeSpan)(endTime - startTime);
            // Use zero for number of searches if the lookup wasn't manual (???)
            string searches = this.scanIdLookup ? "0" : this.searches.ToString();
            this.recordLine[0] = Util.ToMMSS(delta);
            this.recordLine[1] = this.durationHighConfidence ? "High" : "Low";
            this.recordLine[2] = this.endTime.ToString();
            this.recordLine[3] = GetPrettyStringForRecordType(recordType);
            this.recordLine[4] = this.scanIdLookup ? "ID Scan" : "Manual";
            this.recordLine[5] = searches;
            this.recordLine[6] = vbmCancelled ? "Yes" : "No";
            this.recordLine[7] = assistanceRequired ? "Yes" : "No";
            writer.WriteLineArr(recordLine, recordFieldTypes);
        }

        private void ClearState()
        {
            this.state = PollPadState.INIT;
            this.startTime = Util.nullTimestamp;
            this.endTime = Util.nullTimestamp;
            this.vbmCancelled = false;
            this.assistanceRequired = false;
            this.searches = 0;
            this.durationHighConfidence = true;
            this.lastTimestamp = Util.nullTimestamp;
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
            this.writer.WriteLine("Duration (mm:ss)", "Duration Confidence", "End Timestamp", "Event",
                "Lookup Method", "# of Searches", "VBM Cancelled", "Assistance Required");
        }

        public void Done()
        {
        }
    }

}
