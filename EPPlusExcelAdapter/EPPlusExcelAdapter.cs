using ExcelProvider;
using OfficeOpenXml;
using System.IO;

namespace EPPlusExcelAdapter
{
    public class EPPlusExcelAdapter : IExcelAdapter
    {
        private ExcelPackage package { get; set; }
        private ExcelWorksheet selectedSheet { get; set; }

        public void ImportFile(Stream sourceStream)
        {
            package = new ExcelPackage();
            package.Load(sourceStream);
        }

        public void CreateFileForExport()
        {
            package = new ExcelPackage();
        }

        public void AddSheetForExport(string sheetName)
        {
            selectedSheet = package.Workbook.Worksheets.Add(sheetName);
        }

        public void SetValue<T>(int rowNumber, int columnNumber, T dataValue)
        {
            selectedSheet.Cells[rowNumber, columnNumber].Value = dataValue;
        }

        public void SaveExcelFile()
        {
            package.Save();
        }

        public byte[] GetExcelFile()
        {
            return package.GetAsByteArray();
        }
        
        public object GetCellValue(int sheetNumber, int columnIndex, int rowIndex)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetNumber];
            return workSheet.Cells[rowIndex, columnIndex].Value;
        }

        public int GetStartColumn(int sheetNumber)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetNumber];
            return workSheet.Dimension.Start.Column;
        }

        public int GetEndColumn(int sheetNumber)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetNumber];
            return workSheet.Dimension.End.Column;
        }

        public int GetStartRow(int sheetNumber)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetNumber];
            return workSheet.Dimension.Start.Row;
        }

        public int GetEndRow(int sheetNumber)
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets[sheetNumber];
            return workSheet.Dimension.End.Row;
        }
    }
}
