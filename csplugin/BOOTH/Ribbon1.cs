using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace BOOTH
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void ProcessButton_Click(object sender, RibbonControlEventArgs e)
        {
            switch (e.Control.Id)
            {
                case "ProcessSingleButton":
                    Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.VSAP_BMD);
                    //Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.DICE);
                    //Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.DICX);
                    //Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.DS200);
                    //Module1.PollPadProcessing();
                    break;
                case "ProcessAllButton":
                    break;
            }
        }

        private void ImportButton_Click(object sender, RibbonControlEventArgs e)
        {
            switch (e.Control.Id)
            {
                case "ImportDS200Button":
                    Module1.Import_DS200_data();
                    break;
                case "ImportVSAPBMDButton":
                    VSAP_BMD.Import_VSAPBMD_data();
                    break;
                case "ImportDICEButton":
                    Dominion_ICE.Import_DICE_data();
                    break;
                case "ImportDICXButton":
                    Dominion_ICX.Import_DICX_Data();
                    break;
                case "ImportPollPadButton":
                    Module1.PollpadImport();
                    break;
                default:
                    break;
            }
        }

        private void ProcessFolderMenuButton_Click(object sender, RibbonControlEventArgs e)
        {
            switch (e.Control.Id)
            {
                case "ProcessDS200Button":
                    throw new NotImplementedException();
                case "ProcessVSAPBMDButton":
                    Dispatch.ProcessEntireDirectory(LogType.VSAP_BMD);
                    break;
                case "ProcessDICEButton":
                    Dispatch.ProcessEntireDirectory(LogType.DICE);
                    break;
                case "ProcessDICXButton":
                    Dispatch.ProcessEntireDirectory(LogType.DICX);
                    break;
                default:
                    break;
            }
        }

        private void SumStatsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Module1.TestForStat();
        }

        private void TimerOpenButton_Click(object sender, RibbonControlEventArgs e)
        {
            switch (e.Control.Id)
            {
                case "CheckinTimerButton":
                    Timers.LaunchPanelWith(Timers.TimerType.CHECKIN);
                    break;
                case "CheckinArrivalTimerButton":
                    Timers.LaunchPanelWith(Timers.TimerType.CHECKIN_ARRIVAL);
                    break;
                case "VotingBoothTimerButton":
                    Timers.LaunchPanelWith(Timers.TimerType.VOTING_BOOTH);
                    break;
                case "BMDTimerButton":
                    Timers.LaunchPanelWith(Timers.TimerType.BMD);
                    break;
                case "BallotScanningTimerButton":
                    Timers.LaunchPanelWith(Timers.TimerType.BALLOT_SCANNING);
                    break;
                case "ThroughputArrivalTimerButton":
                    Timers.LaunchPanelWith(Timers.TimerType.THROUGHPUT_ARRIVAL);
                    break;
                default:
                    throw new NotImplementedException();
            }
        }
    }
}
