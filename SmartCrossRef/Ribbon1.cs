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

        // Inside your Ribbon code-behind
        private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
        {
            // 1. Get the pane instance for the active window only
            var currentPane = Globals.ThisAddIn.GetTaskPaneInstance();
            if (currentPane != null)
            {
                // 2. Toggle only this specific window's pane
                currentPane.Visible = ((RibbonToggleButton)sender).Checked;

                // 3. Trigger scan if they just opened it
                if (currentPane.Visible)
                {
                    var host = currentPane.Control as CrossRefPaneHostControl;
                    host?.RefreshContextualData();
                }
            }
        }
    }
}
