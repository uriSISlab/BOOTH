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

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void ProcessButton_Click(object sender, RibbonControlEventArgs e)
        {
            switch (e.Control.Id)
            {
                case "ProcessSingleButton":
                    //Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.VSAP_BMD);
                    //Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.DICE);
                    Dispatch.ProcessSheetForLogType(ThisAddIn.app.ActiveWorkbook.ActiveSheet, LogType.DICX);
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
                    Dispatch.processEntireDirectory(LogType.VSAP_BMD);
                    break;
                case "ProcessDICEButton":
                    Dispatch.processEntireDirectory(LogType.DICE);
                    break;
                case "ProcessDICXButton":
                    Dispatch.processEntireDirectory(LogType.DICX);
                    break;
                default:
                    break;
            }
        }
    }
}
