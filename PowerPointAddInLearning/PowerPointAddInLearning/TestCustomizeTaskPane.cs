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
        private bool cancel = false;
        public TestCustomizeTaskPane()
        {
            InitializeComponent();

            Globals.ThisAddIn.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
            Globals.ThisAddIn.Application.PresentationNewSlide += Application_PresentationNewSlide;
            //Globals.ThisAddIn.Application.WindowBeforeDoubleClick += new PowerPoint.Applicatio Application_WindowBeforeDoubleClick;
            //Globals.ThisAddIn.Application.WindowBeforeDoubleClick += new PowerPoint.EApplication_WindowBeforeDoubleClickEventHandler(ApplicationOnWindowBeforeDoubleClick);  //eApplication_WindowBeforeDoubleClickEventHandler
            


        }

        //https://www.itread01.com/content/1550351725.html
        //https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-style-formatting/202
        private void GeneateCalendar(int year, int month, int row, int column, ref Excel.Worksheet ws)//, out int nextRow, out int nextColumn)
        {

            //Day
            //ws.Cells[row + 2, column]
            DateTime firstDate = new DateTime(year, month, 1);
            //firstDate.DayOfWeek
            int endDate = DateTime.DaysInMonth(year, month);

            //Calendar Header
            //ws.Cells[row,column]
            Excel.Range monthHeaderRange = ws.Range[ws.Cells[row, column], ws.Cells[row, column + 6]];
            monthHeaderRange.Merge(true);
            
            monthHeaderRange.Value = firstDate.ToString("MMMM");
            //excelApp.get_Range("A1:A360,B1:E1", Type.Missing).Merge(Type.Missing);

            //ws.get_Range(ws.Cells[1, 1], ws.Cells[1, 2])

            //Week day, Sun-Sat
            ws.Cells[row + 1, column] = "Sun";
            ws.Cells[row + 1, column + 1] = "Mon";
            ws.Cells[row + 1, column + 2] = "Tue";
            ws.Cells[row + 1, column + 3] = "Wed";
            ws.Cells[row + 1, column + 4] = "Thu";
            ws.Cells[row + 1, column + 5] = "Fri";
            ws.Cells[row + 1, column + 6] = "Sat";

            int dateRow = row + 2;
            int dateColumn = column;
            int startedDate = 0;
            switch (firstDate.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    break;
                case DayOfWeek.Monday:
                    startedDate = 1;
                    break;
                case DayOfWeek.Tuesday:
                    startedDate = 2;
                    break;
                case DayOfWeek.Wednesday:
                    startedDate = 3;
                    break;
                case DayOfWeek.Thursday:
                    startedDate = 4;
                    break;
                case DayOfWeek.Friday:
                    startedDate = 5;
                    break;
                case DayOfWeek.Saturday:
                    startedDate = 6;
                    break;
            }

            for (int i = 0; i < endDate; i++)
            {
                //value
                ws.Cells[dateRow + ((i + startedDate) / 7), dateColumn + ((i + startedDate) % 7)] = i + 1;
                //style
            }

            //Excel.Application xl = new Excel.ApplicationClass();

            //Excel.Workbook wb = xl.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorkshe et);

            //Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;

            //ws.Cells[1, 1] = "Testing";

            //Excel.Range range = ws.get_Range(ws.Cells[1, 1], ws.Cells[1, 2]);

            //range.Merge(true);

            //range.Interior.ColorIndex = 36;

            //xl.Visible = true;
        }

        private void Button1_Click(object sender, EventArgs e)
        {

            //string path = Microsoft.SqlServer.Server.MapPath("exportedfiles\\");
            string path = "D:\\";

            if (!Directory.Exists(path))   // CHECK IF THE FOLDER EXISTS. IF NOT, CREATE A NEW FOLDER.
            {
                Directory.CreateDirectory(path);
            }

            File.Delete(path + "laporan2.xlsx"); // DELETE THE FILE BEFORE CREATING A NEW ONE.


            //建立Excel物件
            Excel.Application excelApp = new Excel.Application();

            //新建工作簿
            Excel.Workbook workBook = excelApp.Workbooks.Open(path + "laporan1.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            
            //新建工作表
            Excel.Worksheet ws = workBook.ActiveSheet as Excel.Worksheet;

            //Excel.Range originalRange = ws.Cells[4, 1];
            //originalRange.Copy();
            //Excel.Range destRange = ws.Cells[4, 2];
            //destRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
            //for(int i = 0; i<60; i++)
            //{
            //    //PasteSpecial(4 + i, 1, 4 + i, 2, ref ws);
            //    PasteSpecial(4 + i, 1, ws.Range["E5:E90"], ref ws);

            //    //ws.Range["E5", "F9"]
            //}
            PasteSpecial(ws.Range["A4:A30"], ws.Range["E5:E90"]);
            //PasteSpecial(4, 1, ws.Range["E5:E90"], ref ws);

            //Copies cell C3
            //Range cellToCopy = (Range)activeWorksheet.Cells[3, 3];
            //cellToCopy.Copy();

            ////Paste format only to the cell C5
            //Range cellToPaste = (Range)activeWorksheet.Cells[5, 3];
            //cellToPaste.PasteSpecial(XlPasteType.xlPasteFormats);

            //int year = 2019;
            //int month = 5;
            //int row = 2;
            //int column = 2;
            //GeneateCalendar(2019, 5, 20, 15, ref ws);

            ws.SaveAs(path + "laporan2.xlsx");

            workBook.Save();
            //workBook = null;
            excelApp.Workbooks.Close();
            excelApp.Quit();
            excelApp = null;
            ws = null;
        }
        //private void ApplicationOnWindowBeforeDoubleClick(Selection Sel, ref bool Cancel)
        //{
        //    System.Console.WriteLine("double clck");
        //}
        //private void PasteSpecial(int sourceRow, int sourceColumn, int destinyRow, int destinyColumn, ref Excel.Worksheet ws)
        //{

        //    Excel.Range originalRange = ws.Cells[sourceRow, sourceColumn];
        //    originalRange.Copy();
        //    //Excel.Range
        //    Excel.Range destinyRange = ws.Range["E5", "F9"];
        //    destinyRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
        //}
        //private void PasteSpecial(int sourceRow, int sourceColumn, Excel.Range destinyRange, ref Excel.Worksheet ws)
        //{
        //    Excel.Range originalRange = ws.Cells[sourceRow, sourceColumn];
        //    originalRange.Copy();
        //    destinyRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
        //}
        private void PasteSpecial(int sourceRow, int sourceColumn, Excel.Range destinyRange, ref Excel.Worksheet ws)
        {
            Excel.Range originalRange = ws.Cells[sourceRow, sourceColumn];
            originalRange.Copy();
            destinyRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
        }
        private void PasteSpecial(int sourceRow, int sourceColumn, int destinyRow, int destinyColumn, ref Excel.Worksheet ws)
        {
            Excel.Range originalRange = ws.Cells[sourceRow, sourceColumn];
            originalRange.Copy();
            Excel.Range destinyRange = ws.Cells[destinyRow, destinyColumn];
            destinyRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
        }
        private void PasteSpecial(Excel.Range sourceRange, Excel.Range destinyRange)
        {
            sourceRange.Copy();
            destinyRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats);
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

        private void Button3_Click(object sender, EventArgs e)
        {
            //Microsoft.Office.Interop.PowerPoint.Application app = Globals.ThisAddIn.Application as Microsoft.Office.Interop.PowerPoint.Application;
            //Presentation p = Globals.ThisAddIn.Application.ActivePresentation as Presentation;
            PowerPoint.Slide slide = Globals.ThisAddIn.Application.ActivePresentation.Slides[1];
            string temp = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/embeddedExcel.xlsx";
            //string temp = AppDomain.CurrentDomain.BaseDirectory + "temp.xls";
            //File.Copy(Path, temp, true);
            //Excel.Application xlApp = new Excel.Application();
            //var missing = Type.Missing;
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(temp, missing, missing, missing, missing, missing,
            //    missing, missing, missing, missing, missing, missing, missing, missing, missing);
            //Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
            //string areaA = "Aarea";
            //string areaB = "Barea";
            //string areaC = "Carea";
            //string areaARange = string.Empty;
            //string areaBRange = string.Empty;
            //string areaCRange = string.Empty;

            //foreach (Excel.Name name in xlWorksheet.Names)
            //{
                
            //    if (name.NameLocal.Contains(areaA))
            //    {
            //        areaARange = name.RefersToLocal;
            //    }
            //    else if (name.NameLocal.Contains(areaB))
            //    {
            //        areaBRange = name.RefersToLocal;
            //    }
            //    else if(name.NameLocal.Contains(areaB))
            //    {
            //        areaCRange = name.RefersToLocal;
            //    }
            //}
            //Excel.Range range = xlWorksheet.get_Range(areaARange, missing);
            
            ////range.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

            ////range.Copy(missing);
            ////slide.Design.Application.ActiveWindow.View.Paste();

            ////ShapeRange sr = slide.Shapes.PasteSpecial(
            ////    IconIndex: 0,
            ////    IconLabel: string.Empty,
            ////    IconFileName: string.Empty,
            ////    DisplayAsIcon: Office.MsoTriState.msoFalse,
            ////    DataType:PpPasteDataType.ppPasteDefault, 
            ////    Link: Office.MsoTriState.msoFalse);

            //slide.Shapes[2].ScaleHeight(Factor: 1, Office.MsoTriState.msoCTrue);
            //slide.Shapes[2].ScaleWidth(Factor: 1, Office.MsoTriState.msoCTrue);
            //}
            try
            {
                PowerPoint.Shape shap = slide.Shapes.AddOLEObject(
                    0,
                    0,
                    -1,
                    -1,
                    string.Empty,
                    //ClassName
                    temp,//FileName
                    Office.MsoTriState.msoFalse,
                    temp,//IconFileName
                    1,//IconIndex
                    "test",//IconLabel
                    Office.MsoTriState.msoFalse
                    );
                
                shap.ScaleHeight(Factor: 1, Office.MsoTriState.msoCTrue);
                shap.ScaleWidth(Factor: 1, Office.MsoTriState.msoCTrue);
                slide.Shapes.Range(new int[] { 1, 2 }).ShapeStyle = Office.MsoShapeStyleIndex.msoLineStylePreset11;

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
        private void Button2_Click(object sender, EventArgs e)
        {
            string template = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/embeddedExcel.xlsx";
            string tempFolderPath = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/";

            //var missing = Type.Missing;
            //Load Excel
            //Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(template, 0, false, 5, "", "",
            //        false, Excel.XlPlatform.xlWindows, "",
            //        true, false, 0, true, false, false);

            ////Excel.Application xlApp1 = new Excel.Application();
            //Excel.Workbook wb1 = xlApp.Workbooks.Open(tempFolderPath + "Template1.xlsx", 0, false, 5, "", "",
            //        false, Excel.XlPlatform.xlWindows, "",
            //        true, false, 0, true, false, false);

            //string calendar1 = "Calendar1";
            //string calendar2 = "Calendar2";
            //string calendar3 = "Calendar3";
            //string calendar4 = "Calendar4";

            //string calendar1Area = "";
            //string calendar2Area = "";
            //string calendar3Area = "";
            //string calendar4Area = "";

            //foreach (Excel.Name name in xlWorkbook.Names)
            //{
            //    if (name.NameLocal.Contains(calendar1))
            //    {
            //        calendar1Area = name.RefersToLocal;
            //    }
            //    else if (name.NameLocal.Contains(calendar2))
            //    {
            //        calendar2Area = name.RefersToLocal;
            //    }
            //    else if (name.NameLocal.Contains(calendar3))
            //    {
            //        calendar3Area = name.RefersToLocal;
            //    }
            //    else if (name.NameLocal.Contains(calendar4))
            //    {
            //        calendar4Area = name.RefersToLocal;
            //    }
            //    // Do something to theRange
            //}
            //var calendar1Array = calendar1Area.Substring(1, calendar1Area.Length - 1).Split('!');

            //Excel.Range rangeCal1 = xlWorkbook.Worksheets[calendar1Array[0]].Range(calendar1Array[1], missing);

            //var calendar2Array = calendar2Area.Substring(1, calendar2Area.Length - 1).Split('!');
            //Excel.Range rangeCal2 = xlWorkbook.Worksheets[calendar2Array[0]].Range(calendar2Array[1], missing);

            //var calendar3Array = calendar3Area.Substring(1, calendar3Area.Length - 1).Split('!');
            //Excel.Range rangeCal3 = xlWorkbook.Worksheets[calendar3Array[0]].Range(calendar3Array[1], missing);

            //var calendar4Array = calendar4Area.Substring(1, calendar4Area.Length - 1).Split('!');
            //Excel.Range rangeCal4 = xlWorkbook.Worksheets[calendar4Array[0]].Range(calendar4Array[1], missing);


            //var xlSheetsFrom = xlWorkbook.Sheets as Excel.Sheets;
            //var xlSheetsTo = wb1.Sheets as Excel.Sheets;
            //List<string> nameStr = new List<string>();
            //foreach (Excel.Worksheet sheet in xlSheetsFrom)
            //{
            //    if (sheet.Name.Equals(calendar1Array[0]))
            //    {
            //        Excel.Worksheet curWS = wb1.Worksheets.Add(sheet);
            //        //curWS.Copy(sheet);

            //    }  
            //    //nameStr.Add(sheet.Name);
                
            //    //Excel.Worksheet curWS = wb1.Worksheets.Add(sheet.Name);
            //    //sheet.Copy(wb1.Worksheets);
            //    //wb1.Worksheets.Copy(sheet);

            //    //xlSheetsFrom.Copy(xlSheetsTo);
            //}
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
            //xlWorkbook.Save();
            //xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);

            //wb1.Save();
            //wb1.Close(Type.Missing, Type.Missing, Type.Missing);

            //xlApp.Quit();

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
