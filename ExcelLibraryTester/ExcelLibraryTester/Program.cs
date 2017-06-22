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
            string filePath = @"C:\ExcelTesting\Testdoc.xlsx";
            CellFinderTest(filePath);
            CaptureTest(filePath);
            Console.Read();
        }
        static void CsvTest(string filePath) {
            ExcelReader reader = new ExcelReader(filePath);
            try
            {
                reader.toCSV();
                Console.WriteLine("CSV file made successfully");
            }
            catch (Exception e)
            {
                Console.WriteLine("CSV failed: " + e.Message);
            }
        }
        static void CellFinderTest(string filePath) {
            ExcelReader reader = new ExcelReader(filePath);
            iVector2[] cords = reader.findCells("Sheet1", "Apple");
            foreach (iVector2 cord in cords) {
                Console.WriteLine(cord.ToString());
            }
            Cell[] table = reader.captureCells(cords, "Sheet1");
            foreach (Cell c in table) {
                Console.WriteLine(c.stringValue);
            }
            string[] msgs = reader.printCells(cords, "Sheet1");
            foreach (string s in msgs)
            {
                Console.WriteLine(s);
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
    }
}
