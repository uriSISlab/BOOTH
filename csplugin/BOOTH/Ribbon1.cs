using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
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
                    Worksheet sheet = ThisAddIn.app.ActiveWorkbook.ActiveSheet;
                    ILogProcessor processor = Util.GetCorrectProcessorForImportedSheet(sheet);
                    if (processor != null)
                    {
                        Dispatch.ProcessSheetWithProcessor(sheet, processor);
                    } else
                    {
                        Util.MessageBox("Active worksheet not recognized as a valid log.");
                    }
                    break;
                case "ProcessAllButton":
                    List<string> sheetNames = new List<string>();
                    int processed = 0;
                    foreach (Worksheet sh in ThisAddIn.app.ActiveWorkbook.Sheets)
                    {
                        sheetNames.Add(sh.Name);
                    }
                    foreach (string sheetname in sheetNames)
                    {
                        Worksheet sheet_ = ThisAddIn.app.ActiveWorkbook.Sheets[sheetname];
                        ILogProcessor processor_ = Util.GetCorrectProcessorForImportedSheet(sheet_);
                        if (processor_ != null)
                        {
                            Dispatch.ProcessSheetWithProcessor(sheet_, processor_);
                            processed++;
                        }
                    }
                    Util.MessageBox(processed + " sheet(s) processed.");
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
            TimerBaseForm form;
            string name = ((Microsoft.Office.Tools.Ribbon.RibbonButton)sender).Label;
            Worksheet sheet = Util.TryAddingSheetWithName(name);
            for (int i = 2; i < 50 && sheet == null; i++)
            {
                // Try adding sheets with successively increasing suffixes in case the first name we tried
                // was already taken.
                sheet = Util.TryAddingSheetWithName(name + " " + i);
            }
            if (sheet == null)
            {
                // If sheet is still null after 50 tries to create it, something is seriously wrong. Bail out.
                Util.MessageBox("A worksheet could not be created for the timers.");
                return;
            } 

            switch (e.Control.Id)
            {
                case "CheckinTimerButton":
                    form = TimerBaseForm.CreateForType(TimerBaseForm.TimerFormType.CHECKIN, sheet);
                    break;
                case "CheckinArrivalTimerButton":
                    form = TimerBaseForm.CreateForType(TimerBaseForm.TimerFormType.CHECKIN_ARRIVAL, sheet);
                    break;
                case "VotingBoothTimerButton":
                    form = TimerBaseForm.CreateForType(TimerBaseForm.TimerFormType.VOTING_BOOTH, sheet);
                    break;
                case "BMDTimerButton":
                    form = TimerBaseForm.CreateForType(TimerBaseForm.TimerFormType.BMD, sheet);
                    break;
                case "BallotScanningTimerButton":
                    form = TimerBaseForm.CreateForType(TimerBaseForm.TimerFormType.BALLOT_SCANNING, sheet);
                    break;
                case "ThroughputArrivalTimerButton":
                    form = TimerBaseForm.CreateForType(TimerBaseForm.TimerFormType.THROUGHPUT_ARRIVAL, sheet);
                    break;
                default:
                    throw new NotImplementedException();
            }
            form.Show();
        }

        private void CustomTimersButton_Click(object sender, RibbonControlEventArgs e)
        {
            string name = "Custom timers";
            Worksheet sheet = Util.TryAddingSheetWithName(name);
            for (int i = 2; i < 50 && sheet == null; i++)
            {
                // Try adding sheets with successively increasing suffixes in case the first name we tried
                // was already taken.
                sheet = Util.TryAddingSheetWithName(name + " " + i);
            }
            if (sheet == null)
            {
                // If sheet is still null after 50 tries to create it, something is seriously wrong. Bail out.
                Util.MessageBox("A worksheet could not be created for the timers.");
                return;
            }
            (new CustomTimersCreationForm(sheet)).Show();
        }
    }
}
