using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelManipulator;

namespace ExcelLibraryTester
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"\\acutec.local\Acutec\Network\User Folders\ldawson\Desktop\Excel-Project\NC Documents\Dump.csv";
            CsvParserTest(filePath);
            Console.Read();
        }
        static void CsvTest(string filePath) {
            ExcelReader reader = new ExcelReader(filePath);
            try
            {
                reader.WorksheetToCSV("Dump");
                Console.WriteLine("CSV file made successfully");
            }
            catch (Exception e)
            {
                Console.WriteLine("CSV failed: " + e.Message);
            }
        }
        static void CellFinderTest(string filePath) {
            ExcelReader reader = new ExcelReader(filePath);
            iVector2[] cords = reader.findCells("LOP-053f1", "&&&");
            foreach (iVector2 cord in cords) {
                Console.WriteLine(cord.ToString());
            }
            reader.Close();
        }
        static void CaptureTest(string filePath) {
            ExcelReader reader = new ExcelReader(filePath);
            iVector2[] cords = { new iVector2(1, 2), new iVector2(3, 4), new iVector2(5, 6), new iVector2(3, 2) };
            Cell[] table = reader.captureCells(cords, "Sheet1");
            foreach (Cell c in table)
            {
                Console.WriteLine(c.stringValue);
            }
        }
        static void CsvParserTest(string _filepath) {
            CSVReader csvReader = new CSVReader(_filepath);
            csvReader.Parse();
            string[] values = csvReader.ReturnValues();
            int[] ids = csvReader.ReturnIDs();
            foreach (int i in ids) {
                string s = String.Format("ID:{0} VALUE:{1}", i, values[i]);
                Console.WriteLine(s);
            }
        }
    }
}
