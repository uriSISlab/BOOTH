using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace BOOTH
{
    public static class Timers
    {
        public enum TimerType
        {
            CHECKIN,
            CHECKIN_ARRIVAL,
            BMD,
            VOTING_BOOTH,
            BALLOT_SCANNING,
            THROUGHPUT_ARRIVAL
        };

        public static TimerControl GetTimerControl(TimerType timerType, SheetWriter writer, int number)
        {
            switch (timerType)
            {
                case TimerType.CHECKIN:
                    return new CheckInTimerControl(writer, number);
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
                default:
                    return 5;
            }
        }

        public static void LaunchPanelWith(TimerType timerType)
        {
            TimerBaseForm timerBase = new TimerBaseForm();
            timerBase.PopulateTimers(new TimerType[] { timerType, timerType, timerType, timerType, timerType, timerType },
                ThisAddIn.app.ActiveWorkbook.ActiveSheet);
            timerBase.Show();
        } 
    }
}
