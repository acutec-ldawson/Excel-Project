using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO;
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
            workbook = excel.Workbooks.Open(filePath);
        }
        public void SetFile(string _filePath)
        {
            filePath = _filePath;
        }
        public string GetFile(string _filePath)
        {
            return filePath;
        }
        public void WorkbookToCSV() {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                //Can't access Excel
            }
            else {
                foreach (Worksheet sheet in workbook.Sheets)
                {
                    string csvFile = Path.ChangeExtension(filePath, "_" + sheet.Name + ".csv");
                    sheet.SaveAs(csvFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                }
            }
        }
        public void WorksheetToCSV(string worksheet)
        {
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                //Can't access Excel
            }
            else
            {
                Worksheet sheet = GetSheet(worksheet);
                string csvFile = Path.ChangeExtension(filePath, "_" + sheet.Name + ".csv");
                sheet.SaveAs(csvFile, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
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
                        cellList.Add(new iVector2(cell.Row, cell.Column));
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
    public class Cell {
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
            return String.Format("{0},{1}", x.ToString(), y.ToString());
        }
    }
    public class CSVReader {
        List<int> IDs = new List<int>();
        List<int> AreaIDs = new List<int>();
        List<string> Headers = new List<string>();
        List<string> Groups = new List<string>();
        List<string> Descriptions = new List<string>();
        List<CellValue> Values = new List<CellValue>();
        List<bool> NC = new List<bool>();
        List<bool> CPA = new List<bool>();
        List<string> UniSchema = new List<string>();
        public string filepath;

        public CSVReader(string _filepath) {
            filepath = _filepath;
        }
        public void Parse()
        {
            using (TextFieldParser parser = new TextFieldParser(filepath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.ReadLine();
                while (!parser.EndOfData) {
                    string[] cells = parser.ReadFields();
                    if (cells == null) break;
                    else {
                        int i;
                        Int32.TryParse(cells[0], out i);
                        IDs.Add(i);
                        /*string s = cells[1];
                        char id = s[0];
                        int j = 0;
                        Int32.TryParse(id.ToString(), out j);
                        AreaIDs.Add(j);*/
                        Headers.Add(cells[2]);
                        Groups.Add(cells[3]);
                        Descriptions.Add(cells[4]);
                        Values.Add(new CellValue(cells[6],GetDesiredType(cells[5])));
                        NC.Add(StringToBool(cells[7]));
                        NC.Add(StringToBool(cells[8]));
                        UniSchema.Add(cells[9]);
                    }
                }
            }
        }
        private CellValue.ValueType GetDesiredType(string s) {
            switch (s) {
                case "Boolean":
                    return CellValue.ValueType.Boolean;
                case "Float":
                    return CellValue.ValueType.Float;
                case "Integer":
                    return CellValue.ValueType.Integer;
                case "Date":
                    return CellValue.ValueType.Date;
                default:
                    return CellValue.ValueType.String;
            }
        }
        private bool StringToBool(string s) {
            switch (s)
            {
                case "TRUE":
                    return true;
                default:
                    return false;
            }
        }
        public int[] ReturnIDs() {
            return IDs.ToArray();
        }
        public string[] ReturnValues() {
            List<string> sValues = new List<string>();
            foreach (CellValue val in Values) {
                sValues.Add(val.ToString());
            }
            return sValues.ToArray();
        }
    }
    public struct Date {
        public int day, month, year;
        public Date(int _day, int _month, int _year) {
            day = _day;
            month = _month;
            year = _year;
        }
    }
    public struct CellValue {
        public enum ValueType {
            Date,
            Integer,
            String,
            Boolean,
            Float
        }
        public ValueType DesiredType { get; }
        public Type type { get; }
        public object value {get;}
        public string sValue;
        public CellValue(object _value, ValueType _DesiredType)
        {
            sValue = (string)_value;
            DesiredType = _DesiredType;
            if (DesiredType == ValueType.String)
            {
                value = _value;
                type = typeof(string);
            }
            else if (DesiredType == ValueType.Integer)
            {
                int i;
                Int32.TryParse((string)_value, out i);
                value = i;
                type = typeof(int);
            }
            else if (DesiredType == ValueType.Date)
            {
                value = _value;
                type = typeof(Date);
                char[] delimiters = { '/', '/' };
                string[] vars = value.ToString().Split(delimiters);
                int month = Int32.Parse(vars[0]);
                int day = Int32.Parse(vars[1]);
                int year = Int32.Parse(vars[2]);
                value = new Date(_day: day, _month: month, _year: year);
            }
            else if (DesiredType == ValueType.Boolean)
            {
                value = _value;
                type = typeof(bool);
                if (_value.GetType().Equals(typeof(string)))
                {
                    if ((string)_value == "TRUE" || (string)_value == "T" || (string)_value == "True" || (string)_value == "true") value = true;
                    if ((string)_value == "FALSE" || (string)_value == "F" || (string)_value == "False" || (string)_value == "false") value = false;
                }
            }
            else if (DesiredType == ValueType.Float)
            {
                value = _value;
                float f;
                float.TryParse((string)_value, out f);
                value = f;
                type = typeof(float);
            }
            else
            {
                value = _value;
                type = typeof(string);
            }
        }
        public override string ToString() {
            return sValue;
        }
    }
}
