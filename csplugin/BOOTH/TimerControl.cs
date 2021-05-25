using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOOTH
{
    public class TimerControl : UserControl
    {
        public enum TimerType
        {
            CHECKIN,
            ARRIVAL,
            VOTING_BOOTH,
            BMD,
            BALLOT_SCANNING,
            THROUGHPUT,
        };

        public static TimerType[] GetMainPanelTimerTypes()
        {
            return new TimerType[] { TimerType.CHECKIN, TimerType.VOTING_BOOTH, TimerType.BMD,
                TimerType.BALLOT_SCANNING, TimerType.THROUGHPUT };
        }

        public static string GetNiceNameForTimerType(TimerType timerType)
        {
            switch (timerType)
            {
                case TimerType.CHECKIN:
                    return "Check in";
                case TimerType.ARRIVAL:
                    return "Arrival";
                case TimerType.VOTING_BOOTH:
                    return "Voting Booth";
                case TimerType.BALLOT_SCANNING:
                    return "Ballot Scanning";
                case TimerType.THROUGHPUT:
                    return "Throughput";
                case TimerType.BMD:
                    return "BMD";
                default:
                    return "Unknown";
            }
        }

        public static int GetColumnCountForTimerType(TimerType timerType)
        {
            switch (timerType)
            {
                case TimerType.CHECKIN:
                    return 5;
                case TimerType.ARRIVAL:
                    return 3;
                case TimerType.VOTING_BOOTH:
                case TimerType.BALLOT_SCANNING:
                case TimerType.THROUGHPUT:
                    return 4;
                case TimerType.BMD:
                default:
                    return 5;
            }
        }
        public static TimerControl GetTimerControl(TimerType timerType, SheetWriter writer, int number)
        {
            switch (timerType)
            {
                case TimerType.CHECKIN:
                    return new CheckInTimerControl(writer, number);
                case TimerType.ARRIVAL:
                    return new ArrivalTimerControl(writer);
                case TimerType.VOTING_BOOTH:
                    return new VotingBoothTimerControl(writer, number);
                case TimerType.BMD:
                    return new BMDTimerControl(writer, number);
                case TimerType.BALLOT_SCANNING:
                    return new BallotScanningTimerControl(writer, number);
                case TimerType.THROUGHPUT:
                    return new ThroughputTimerControl(writer, number);
                default:
                    return null;
            }
        }

        protected static string[] GetShortCutsForStartAndStop(int number)
        {
            string startShortcut, stopShortcut;
            if (number < 5)
            {
                startShortcut = Convert.ToString(number * 2 - 1);
                stopShortcut = Convert.ToString(number * 2);
            } else if (number == 5)
            {
                startShortcut = Convert.ToString(9);
                stopShortcut = Convert.ToString(0);
            } else if (number == 6)
            {
                startShortcut = "-";
                stopShortcut = "=";
            } else
            {
                startShortcut = null;
                stopShortcut = null;
            }
            return new string[] { startShortcut, stopShortcut };
        }

        protected SheetWriter writer;
        protected int number; 

        public TimerControl()
        {
            // NOTE: This constructor is here because the UI designer needs the base class
            // of a UI element to be non-abstract and to have a no-argument constructor.
            // This constructor should not be used in practice and this class should be treated
            // as abstract.
        }

        public TimerControl(SheetWriter writer, int number)
        {
            this.writer = writer;
            this.number = number;
        }

        public virtual string GetHeadingText()
        {
            throw new NotImplementedException();
        }

        public virtual void AddComment(string comment)
        {
            throw new NotImplementedException();
        }
    }
}
