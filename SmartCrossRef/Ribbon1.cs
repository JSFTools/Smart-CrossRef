using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartCrossRef
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
        {
            bool isPressed = ((RibbonToggleButton)sender).Checked;

            // Call the updated master synchronization method 
            Globals.ThisAddIn.SyncAllPanesVisibility(isPressed);
        }
    }
}
