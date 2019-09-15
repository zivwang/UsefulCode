using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Ribbon;

namespace PowerPointAddInLearning
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        
        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {

            //TestCustomizeTaskPane testCustomizeTaskPane = new TestCustomizeTaskPane(app.ActivePresentation);

            //CustomTaskPane ctp = this.myCustomTaskPanes.Add(testCustomizeTaskPane, "My Task Pane");
            //ctp.Visible = true;

        }
    }
}
