using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApp1
{
    class CSTest
    {
        public void Run()
        {
            string template = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/embeddedExcel.xlsx";
            string tempFolderPath = "E:/Dev/OfficeAddInLearning/PowerPointAddInLearning/";

            var missing = Type.Missing;
            //Load Excel
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(template);

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

            Excel.Application newXlAppCal1 = new Excel.Application();
            Excel.Workbook xlWBCal1 = newXlAppCal1.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet wsCal1 = xlWBCal1.Worksheets[1];
            wsCal1.Name = calendar1Array[0];
            //cleanup  

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //xlWorkbook.Worksheets[calendar1Array[0]]

            //Excel.Range rangeCal1Dest = xlWBCal1.Worksheets[calendar1Array[0]].Range(calendar1Array[1], missing);
            //rangeCal1.Copy(rangeCal1Dest);

            Excel.Application xlApp1 = new Excel.Application();
            Excel.Workbook wb1 = xlApp1.Workbooks.Open(tempFolderPath + "Template4.xlsx");//Excel.Workbooks.Open(tempFolderPath + "Template4.xlsx", missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
            Microsoft.Office.Interop.Excel.Worksheet wx1 = wb1.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
            wx1.Copy(xlWorkbook.Worksheets[calendar1Array[0]]);

            wb1.Save();
            xlApp1.Quit();
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

            xlWBCal1.Close();
            newXlAppCal1.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wsCal1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWBCal1);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(newXlAppCal1);
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
    }
}
