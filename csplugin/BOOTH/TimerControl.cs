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
