using ExcelProvider;
using System;
using System.Collections.Generic;
using System.IO;
using EP = EPPlusExcelAdapter;

namespace Client
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Give path and file name to save excel file (Sample : C:\\Trial.xlsx) : ");
            var path = Console.ReadLine();

            IExcelAdapter adapter = new EP.EPPlusExcelAdapter(); //Resolve with IOC
            IExcelExporter exporter = new DefaultExcelExporter(); //Resolve with IOC

            var data = GetData();
            var targetSourceMap = new Dictionary<string, string> //For saving as other colum names
            {
                { "Column1", "Column1" },
                { "Column2", "Column2" },
                { "Column3", "Column3" }
            };

            adapter.CreateFileForExport();
            exporter.ExportListAsSheet<Trial>(adapter, data, "Trials", targetSourceMap, true);
            exporter.ExportListAsSheet<Trial>(adapter, data, "Trials2", targetSourceMap, true);
            var fileBytes = exporter.GetExcelFile(adapter);

            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                fs.Write(fileBytes, 0, fileBytes.Length);
                fs.Flush();
            }

            Console.WriteLine("Press any key to import same file to list.");
            Console.ReadKey();

            adapter = new EP.EPPlusExcelAdapter();
            IExcelImporter importer = new DefaultExcelImporter();

            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                adapter.ImportFile(fs);
                var endRow = adapter.GetEndRow(1);
                var result = importer.Populate<Trial>(adapter, 1, targetSourceMap, 1, endRow, OnRowImportError);
            }
        }

        public static void OnRowImportError(Trial data, string columnName, Exception ex)
        {

        }

        static List<Trial> GetData()
        {
            var result = new List<Trial>
            {
                new Trial
                {
                    Column1 = "Value1", Column2 = 5, Column3 = 100.100f
                },
                new Trial
                {
                    Column1 = "Value2", Column2 = 10, Column3 = 200.200f
                },
                new Trial
                {
                    Column1 = "Value3", Column2 = 15, Column3 = 300.300f
                },
                new Trial
                {
                    Column1 = "Value4", Column2 = 20, Column3 = 400.400f
                },
                new Trial
                {
                    Column1 = "Value5", Column2 = 25, Column3 = 500.500f
                }
            };

            return result;
        }
    }


    public class Trial
    {
        public Trial()
        {

        }

        public string Column1 { get; set; }

        public int Column2 { get; set; }

        public float Column3 { get; set; }
    }
}
