using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace BOOTH
{
    public static class Timers
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

        public enum TimerFormType
        {
            CHECKIN,
            CHECKIN_ARRIVAL,
            VOTING_BOOTH,
            BMD,
            BALLOT_SCANNING,
            THROUGHPUT_ARRIVAL
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

        public static void LaunchPanelWith(TimerFormType timerFormType)
        {
            TimerBaseForm timerBase = TimerBaseForm.CreateForType(timerFormType);
            timerBase.Show();
        } 
    }
}
