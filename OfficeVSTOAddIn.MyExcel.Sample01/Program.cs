using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeVSTOAddIn.MyExcel.Sample01
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = Path.Combine(Environment.CurrentDirectory, DateTime.Now.ToString("yyyyMMddHHmmssfff") + "-test.xlsx");
            Console.WriteLine(filePath);
            CreateExcelFile(filePath);
        }

        static void CreateExcelFile(string filePath)
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet xlWorksheet = xlWorkbook.Worksheets[1];

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= 10; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    xlWorksheet.Cells[i, j] = i + " " + j;
                }
            }

            xlWorkbook.SaveCopyAs(filePath);

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //close and release
            xlWorkbook.Close(false);
            //quit and release
            xlApp.Quit();

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlWorkbook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
