using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using MouseKeyboardActivityMonitor;
using MouseKeyboardActivityMonitor.WinApi;
using System.Diagnostics;
using System.Windows.Forms;

namespace PowerPointAddInLearning
{
    public partial class ThisAddIn
    {
        
        private Microsoft.Office.Tools.CustomTaskPaneCollection myCustomTaskPanes;
        private PowerPoint.Presentation currentPresentation;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Application.WindowSelectionChange += Application_WindowSelectionChange;
            //this.Application.PresentationNewSlide +=new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
            //Application.ActivePresentation = new PowerPoint.Presentation
            //PowerPoint.Presentation pre = new 
            Application.WindowBeforeDoubleClick += new PowerPoint.EApplication_WindowBeforeDoubleClickEventHandler(ApplicationOnWindowBeforeDoubleClick);  //eApplication_WindowBeforeDoubleClickEventHandler

            //Application.WindowBeforeDoubleClick += new PowerPoint.EApplication_WindowBeforeDoubleClickEventHandler(ApplicationOnWindowBeforeDoubleClick);
            MouseHookListener mouseHookListener = new MouseHookListener(new AppHooker()) { Enabled = true };
            mouseHookListener.MouseDoubleClick += MouseHookListener_MouseDoubleClick;
            mouseHookListener.MouseClickExt += MouseHookListener_MouseClickExt;
        }

        private void MouseHookListener_MouseClickExt(object sender, MouseEventExtArgs e)
        {
            //throw new NotImplementedException();
            Debug.Print("mouseClickExt");
            if (e.Clicks == 2)
            {
                //TODO Insert your double-click code here
                e.Handled = true;
                Debug.Print("mouseClickExt Handled");
            }
            
        }

        private void MouseHookListener_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            MessageBox.Show("Sorry, the object cant be double clicked");

            //throw new NotImplementedException();
            Debug.Print("MouseHookListener_MouseDoubleClick");
            
        }

        private void ApplicationOnWindowBeforeDoubleClick(PowerPoint.Selection Sel, ref bool Cancel)
        {
            System.Console.WriteLine("double clck");
            //throw new NotImplementedException();
        }


        //private void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        //{
        //    //Sld.Shapes.AddOLEObject("")
        //    //throw new NotImplementedException();
        //    PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
        //    textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        //}


        //private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        //{
        //    //throw new NotImplementedException();
        //    //if (Sel.SlideRange.End - Sel.Range.Start > 10)
        //    //{
        //    //    Word.Range range = Sel.Range;

        //    //    System.Diagnostics.Debug.WriteLine("Range Start Position: {0}  Range End Position: {1}", range.Start, range.End);

        //    //    range.Bookmarks.Add("MyBookmark");
        //    //}
        //}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion



        // TODO:  Follow these steps to enable the Ribbon (XML) item:

        // 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            //currentPresentation = Application.ActivePresentation as PowerPoint.Presentation;
            myCustomTaskPanes = this.CustomTaskPanes as Microsoft.Office.Tools.CustomTaskPaneCollection;
            PowerPoint.Application app =  this.Application as Microsoft.Office.Interop.PowerPoint.Application;
            return new Ribbon();
        }

        // 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
        //    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
        //    move your code from the event handlers to the callback methods and modify the code to work with the
        //    Ribbon extensibility (RibbonX) programming model.

        // 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

        // For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.

    }
}
