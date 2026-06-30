using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace OutlookClassifierAddIn5.UI
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnTogglePane_Click(object sender, RibbonControlEventArgs e)
            => Globals.ThisAddIn.TogglePane();


        private async void btnRetrain_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.PaneControl == null) Globals.ThisAddIn.TogglePane(); // creates & shows
            if (Globals.ThisAddIn.PaneControl != null)
                await Globals.ThisAddIn.PaneControl.TrainModelAsync(savePath: null);
        }

        private void btnUndo_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.PaneControl == null) Globals.ThisAddIn.TogglePane();
            Globals.ThisAddIn?.PaneControl?.TryUndoLastMove();
        }

        private void chkAutoApprove_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.PaneControl == null) Globals.ThisAddIn.TogglePane();
            var check = (RibbonCheckBox)sender;
            var enabled = Globals.ThisAddIn?.PaneControl?.SetAutoApprove(check.Checked) ?? false;
            check.Checked = enabled;
        }
    }
}
