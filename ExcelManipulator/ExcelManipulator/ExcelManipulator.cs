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
        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        Microsoft.Office.Interop.Excel.Workbook workbook = null;
        string filePath = "";
        public ExcelReader(string _filePath) {
            filePath = _filePath;
            setDisplay();
            workbook = excel.Workbooks.Open(filePath, ReadOnly: true);
        }
        public void SetFile(string _filePath)
        {
            filePath = _filePath;
        }
        public string GetFile(string _filePath)
        {
            return filePath;
        }
        public void toCSV() {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                //Can't access Excel
            }
            else {
                
                foreach (Worksheet sheet in workbook.Sheets)
                {
                    string csvFile = Path.ChangeExtension(filePath, "_"+sheet.Name+".csv");
                    sheet.SaveAs(csvFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }
            }
        }
        /*
         *  
         */
        /// <summary>
        /// Searches the filepath excel document and returns a list of all the cells that contained the flag string
        /// This allows the user to quickly develope a list of all the cells that the user will need to pull values from
        /// </summary>
        /// <param name="worksheet">The worksheet that you are searching</param>
        /// <param name="flag">The string that the program will search the worksheet for</param>
        /// <returns></returns>
        public iVector2[] findCells(string worksheet, string flag) {
            List<iVector2> cellList = new List<iVector2>();
            Worksheet sheet = GetSheet(worksheet);

            foreach (Range cell in sheet.UsedRange) {
                try
                {
                    var val = (string)cell.Value;
                    if (val == flag)
                    {
                        cellList.Add(new iVector2(cell.Row,cell.Column));
                    }
                }
                catch (Exception e) { }
            }
            return cellList.ToArray();
        }
        private Worksheet GetSheet(string sheetName) {
            Worksheet get = null;

            foreach (Worksheet sheet in workbook.Sheets) {
                if (sheet.Name == sheetName) {
                    get = sheet;
                }
            }
            try
            {
                return get;
            }
            finally {
            }
        }
        public Cell[] captureCells(iVector2[] coordinates, string worksheet) {
            List<Cell> Table = new List<Cell>();
            Worksheet sheet = GetSheet(worksheet);

            foreach (iVector2 dir in coordinates) {
                Range r = sheet.UsedRange.Cells[dir.x, dir.y];
                Cell c = new Cell(dir, r.Value);
                Table.Add(c);
            }
            return Table.ToArray();
        }
        public string[] printCells(iVector2[] coordinates, string worksheet) {
            List<string> Stuff = new List<string>();
            Worksheet sheet = GetSheet(worksheet);

            foreach (iVector2 dir in coordinates)
            {
                Range r = sheet.UsedRange.Cells[dir.x, dir.y];
                Cell c = new Cell(dir, r.Value);
                Stuff.Add(c.stringValue);
            }
            return Stuff.ToArray();
        }
        public void Close() {
            workbook.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
        public void setDisplay() {
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.EnableEvents = false;
        }
    }
    public class Cell{
        public object value;
        public string stringValue;
        public iVector2 coordinate;
        Type type;
        public Cell(iVector2 dir, object val) {
            coordinate = dir;
            value = val;
            stringValue = val.ToString();
            type = val.GetType();
        }
        public Cell(int i, int j, object val)
        {
            coordinate = new iVector2(i, j);
            value = val; 
            stringValue = val.ToString();
            type = val.GetType();
        }
    }
    public class iVector2 {
        //Row
        public int x;
        //Column
        public int y;
        public iVector2(int i, int j) {
            x = i;
            y = j;
        }
        public override string ToString()
        {
            return String.Format("{0},{1}",x.ToString(),y.ToString());
        }
    }
}
