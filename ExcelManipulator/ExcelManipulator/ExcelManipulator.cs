using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelManipulator
{
    public class ExcelReader
    {

        string filePath = "";
        public ExcelReader() {
        }
        public ExcelReader(string _filePath) {
            filePath = _filePath;
        }
        public void SetFile(string _filePath)
        {
            filePath = _filePath;
        }
        public void toCSV() {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                //Can't access Excel
            }
            else {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                app.DisplayAlerts = false;

                Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Open(filePath);
                foreach (Worksheet sheet in workbook.Sheets)
                {
                    string csvFile = Path.ChangeExtension(filePath, "_"+sheet.Name+".csv");
                    sheet.SaveAs(csvFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                }

                workbook.Close();
                app.Quit();
            }
        }


    }
}
