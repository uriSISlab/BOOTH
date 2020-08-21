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
    }
}
