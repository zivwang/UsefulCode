using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PowerPointAddInLearning
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private readonly string taskPane = "Custom Task Pane";
        //private Microsoft.Office.Tools.CustomTaskPaneCollection myCustomTaskPanes;
        //private Application app;

        public Ribbon()
        {

        }
        //public Ribbon(Microsoft.Office.Tools.CustomTaskPaneCollection myCustomTaskPanes, Application application)
        //{
        //    this.myCustomTaskPanes = myCustomTaskPanes;
        //    this.app = application;
        //}
        //public Ribbon()
        //{
        //}



        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointAddInLearning.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {

            //TestCustomizeTaskPane testCustomizeTaskPane = new TestCustomizeTaskPane(app.ActivePresentation);

            //CustomTaskPane ctp = this.myCustomTaskPanes.Add(testCustomizeTaskPane, "My Task Pane");
            //ctp.Visible = true;

        }

        public void Button1_Click(Office.IRibbonControl control)
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            {

                if (!Globals.ThisAddIn.CustomTaskPanes.Any(i => i.Title == taskPane))
                {
                    Globals.ThisAddIn.CustomTaskPanes.RemoveAt(0);
                    TestCustomizeTaskPane testCustomizeTaskPane = new TestCustomizeTaskPane();
                    Microsoft.Office.Tools.CustomTaskPane ctp = Globals.ThisAddIn.CustomTaskPanes.Add(testCustomizeTaskPane, taskPane);
                    ctp.Visible = true;
                }
                else
                {
                    Globals.ThisAddIn.CustomTaskPanes.RemoveAt(0);
                }
                //Globals.ThisAddIn.Application.ActivePresentation
                //Word.Range currentRange = Globals.ThisAddIn.Application.Selection.Range;
                //currentRange.Text = "This text was added by the Ribbon.";
            }
            else
            {
                TestCustomizeTaskPane testCustomizeTaskPane = new TestCustomizeTaskPane();
                Microsoft.Office.Tools.CustomTaskPane ctp = Globals.ThisAddIn.CustomTaskPanes.Add(testCustomizeTaskPane, taskPane);
                ctp.Visible = true;
            }
        }
    }
}
