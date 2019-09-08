using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
//using Office = Microsoft.Office.Core;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace PowerPointAddInLearning
{
    public partial class TestCustomizeTaskPane : UserControl
    {
        //private Microsoft.Office.Interop.PowerPoint.Presentation cp;
        private Slide currentSld;
        private Shape textBox;

        public TestCustomizeTaskPane()
        {
            InitializeComponent();

            Globals.ThisAddIn.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            Globals.ThisAddIn.Application.PresentationNewSlide += Application_PresentationNewSlide;
        }

        private void Application_PresentationNewSlide(Slide Sld)
        {
            currentSld = Sld;
           
        }

        private void Application_SlideSelectionChanged(SlideRange sldRange)
        {
            //sldRange.Select();
            //throw new NotImplementedException();
            //Console.WriteLine(SldRange[1].SlideNumber);


        }

        //public void LoadEmbeddedExcel()
        //{

        //}

        private void Button2_Click(object sender, EventArgs e)
        {



            //Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application as Microsoft.Office.Interop.PowerPoint.Application;
            //Presentation p = Globals.ThisAddIn.Application.ActivePresentation as Presentation;
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActivePresentation.Slides[1];

            string temp = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/embeddedExcel.xlsx";
            //string temp = AppDomain.CurrentDomain.BaseDirectory + "temp.xls";
            //File.Copy(Path, temp, true);
            Excel.Application xlApp = new Excel.Application();
            var missing = Type.Missing;
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(temp, missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
            //var namedRanges = xlWorkbook.Names.Item.;//.Add(Name: "MyTable", RefersToR1C1: "=Sheet1!R1C1:R1C5");
            //Excel.Range destRange = xlWorksheet.get_Range("Aarea");
            string areaA = "Aarea";
            string areaB = "Barea";
            string areaC = "Carea";
            string areaARange = string.Empty;
            string areaBRange = string.Empty;
            string areaCRange = string.Empty;

            foreach (Excel.Name name in xlWorksheet.Names)
            {
                
                if (name.NameLocal.Contains(areaA))
                {
                    areaARange = name.RefersToLocal;
                }
                else if (name.NameLocal.Contains(areaB))
                {
                    areaBRange = name.RefersToLocal;
                }
                else if(name.NameLocal.Contains(areaB))
                {
                    areaCRange = name.RefersToLocal;
                }

        // Do something to theRange
            }
            Excel.Range range = xlWorksheet.get_Range(areaARange, missing);
            //range.Copy()


            //range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

            //range.Copy(missing);
            //slide.Design.Application.ActiveWindow.View.Paste();

            //ShapeRange sr = slide.Shapes.PasteSpecial(
            //    IconIndex: 0,
            //    IconLabel: string.Empty,
            //    IconFileName: string.Empty,
            //    DisplayAsIcon: Office.MsoTriState.msoFalse,
            //    DataType:PpPasteDataType.ppPasteDefault, 
            //    Link: Office.MsoTriState.msoFalse);

            //slide.Shapes[2].ScaleHeight(Factor: 1, Office.MsoTriState.msoCTrue);
            //slide.Shapes[2].ScaleWidth(Factor: 1, Office.MsoTriState.msoCTrue);
            //}
            try
            {
                //PowerPoint.Shape shap = slide.Shapes.AddOLEObject(
                //    0,
                //    0,
                //    -1,
                //    -1,
                //    string.Empty,
                //    //ClassName
                //    temp,//FileName
                //    Office.MsoTriState.msoFalse,
                //    temp,//IconFileName
                //    1,//IconIndex
                //    "test",//IconLabel
                //    Office.MsoTriState.msoFalse
                //    );
                //shap.
                //shap.ScaleHeight(Factor: 1, Office.MsoTriState.msoCTrue);
                //shap.ScaleWidth(Factor: 1, Office.MsoTriState.msoCTrue);
                //slide.Shapes.Range(new int[] { 1, 2 }).ShapeStyle = Office.MsoShapeStyleIndex.msoLineStylePreset11;

                //shap.

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                //File.Delete(temp);
                //Close();
            }

        }

        //get config (startMonth, region, publicHoliday
        DateTime startMonth = new DateTime(2019, 09, 01);
        string region;
        string publicHoliday;
        bool showCalendar;
        int noOfCalendarDisplay; //1/2/3/4
        bool showTimeLine;

        //get excel template base on config
        private void Button1_Click(object sender, EventArgs e)
        {
            string template = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/embeddedExcel.xlsx";
            string tempFolderPath = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/";

            var missing = Type.Missing;
            //Load Excel
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(template, 0, false, 5, "", "",
                    false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

            //Excel.Application xlApp1 = new Excel.Application();
            Excel.Workbook wb1 = xlApp.Workbooks.Open(tempFolderPath + "Template1.xlsx", 0, false, 5, "", "",
                    false, Excel.XlPlatform.xlWindows, "",
                    true, false, 0, true, false, false);

            string calendar1 = "Calendar1";
            string calendar2 = "Calendar2";
            string calendar3 = "Calendar3";
            string calendar4 = "Calendar4";

            string calendar1Area = "";
            string calendar2Area = "";
            string calendar3Area = "";
            string calendar4Area = "";

            foreach (Excel.Name name in xlWorkbook.Names)
            {
                if (name.NameLocal.Contains(calendar1))
                {
                    calendar1Area = name.RefersToLocal;
                }
                else if (name.NameLocal.Contains(calendar2))
                {
                    calendar2Area = name.RefersToLocal;
                }
                else if (name.NameLocal.Contains(calendar3))
                {
                    calendar3Area = name.RefersToLocal;
                }
                else if (name.NameLocal.Contains(calendar4))
                {
                    calendar4Area = name.RefersToLocal;
                }
                // Do something to theRange
            }
            var calendar1Array = calendar1Area.Substring(1, calendar1Area.Length - 1).Split('!');

            Excel.Range rangeCal1 = xlWorkbook.Worksheets[calendar1Array[0]].Range(calendar1Array[1], missing);

            var calendar2Array = calendar2Area.Substring(1, calendar2Area.Length - 1).Split('!');
            Excel.Range rangeCal2 = xlWorkbook.Worksheets[calendar2Array[0]].Range(calendar2Array[1], missing);

            var calendar3Array = calendar3Area.Substring(1, calendar3Area.Length - 1).Split('!');
            Excel.Range rangeCal3 = xlWorkbook.Worksheets[calendar3Array[0]].Range(calendar3Array[1], missing);

            var calendar4Array = calendar4Area.Substring(1, calendar4Area.Length - 1).Split('!');
            Excel.Range rangeCal4 = xlWorkbook.Worksheets[calendar4Array[0]].Range(calendar4Array[1], missing);


            var xlSheetsFrom = xlWorkbook.Sheets as Excel.Sheets;
            var xlSheetsTo = wb1.Sheets as Excel.Sheets;
            List<string> nameStr = new List<string>();
            foreach (Excel.Worksheet sheet in xlSheetsFrom)
            {
                if (sheet.Name.Equals(calendar1Array[0]))
                {
                    Excel.Worksheet curWS = wb1.Worksheets.Add(sheet);
                    //curWS.Copy(sheet);

                }  
                //nameStr.Add(sheet.Name);
                
                //Excel.Worksheet curWS = wb1.Worksheets.Add(sheet.Name);
                //sheet.Copy(wb1.Worksheets);
                //wb1.Worksheets.Copy(sheet);

                //xlSheetsFrom.Copy(xlSheetsTo);
            }
            //Excel.Application newXlAppCal1 = new Excel.Application();
            //Excel.Workbook xlWBCal1 = newXlAppCal1.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            //Excel.Worksheet wsCal1 = xlWBCal1.Worksheets[1];
            //wsCal1.Name = calendar1Array[0];
            //cleanup  

            //GC.Collect();
            //GC.WaitForPendingFinalizers();

            //Excel.Range rangeCal1Dest = xlWBCal1.Worksheets[calendar1Array[0]].Range(calendar1Array[1], missing);
            //rangeCal1.Copy(rangeCal1Dest);

            //string tempExel1 = tempFolderPath + calendar1Array[0].ToString() + ".xlsx";
            //xlWBCal1.SaveAs(tempExel1, 
            //    Excel.XlFileFormat.xlOpenXMLWorkbook, 
            //    missing,
            //    missing,
            //    missing,
            //    missing, 
            //    Excel.XlSaveAsAccessMode.xlExclusive,
            //    missing,
            //    missing,
            //    missing,
            //    missing,
            //    missing);

            //Excel.Workbooks.Open(tempFolderPath + "Template4.xlsx", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //Microsoft.Office.Interop.Excel.Worksheet wx1 = wb1.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            //wx1.Copy(xlWorkbook.Worksheets[calendar1Array[0]]);

            //xlSheetsTo.Add(xlWorkbook.Worksheets[calendar1Array[0]], Type.Missing, Type.Missing, Type.Missing);
            //wsCal1.Copy(wx1);
            xlWorkbook.Save();
            xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);

            wb1.Save();
            wb1.Close(Type.Missing, Type.Missing, Type.Missing);

            xlApp.Quit();

            //xlWBCal1.Close();
            //newXlAppCal1.Quit();
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCal1);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWBCal1);
            //System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlAppCal1);
            //Excel.Application newXlAppCal2 = new Microsoft.Office.Interop.Excel.Application();
            //Excel.Workbook xlWBCal2 = newXlAppCal2.Workbooks.Add(missing);
            //var wsCal2 = xlWBCal2.Worksheets.get_Item(1);
            //wsCal2.Name = calendar2Array[0];
            //Excel.Range rangeCal2Dest = xlWBCal2.Worksheets[calendar2Array[0]].Range(calendar2Array[1], missing);
            //rangeCal2.Copy(rangeCal2Dest);
            //xlWBCal2.SaveAs(tempFolderPath + calendar2Array[0] + ".xlsx");

            //Excel.Application newXlAppCal3 = new Microsoft.Office.Interop.Excel.Application();
            //Excel.Workbook xlWBCal3 = newXlAppCal3.Workbooks.Add(missing);
            //var wsCal3 = xlWBCal3.Worksheets.get_Item(1);
            //wsCal3.Name = calendar3Array[0];
            //Excel.Range rangeCal3Dest = xlWBCal3.Worksheets[calendar3Array[0]].Range(calendar3Array[1], missing);
            //rangeCal3.Copy(rangeCal3Dest);
            //xlWBCal3.SaveAs(tempFolderPath + calendar3Array[0] + ".xlsx");

            //Excel.Application newXlAppCal4 = new Microsoft.Office.Interop.Excel.Application();
            //Excel.Workbook xlWBCal4 = newXlAppCal4.Workbooks.Add(missing);
            //var wsCal4 = xlWBCal4.Worksheets.get_Item(1);
            //wsCal4.Name = calendar4Array[0];
            //Excel.Range rangeCal4Dest = xlWBCal4.Worksheets[calendar4Array[0]].Range(calendar4Array[1], missing);
            //rangeCal4.Copy(rangeCal4Dest);
            //xlWBCal4.SaveAs(tempFolderPath + calendar4Array[0] + ".xlsx");
        }
        //copy range to temp file.

        //generate calendar/timeline/

        //embed OLEObject to ppt;
    }
}
