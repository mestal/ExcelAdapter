using System.IO;

namespace ExcelProvider
{
    public interface IExcelAdapter
    {
        void ImportFile(Stream sourceStream);
        
        object GetCellValue(int sheetNumber, int columnIndex, int rowIndex);

        int GetStartColumn(int sheetNumber);

        int GetEndColumn(int sheetNumber);

        int GetStartRow(int sheetNumber);

        int GetEndRow(int sheetNumber);

        void CreateFileForExport();

        void AddSheetForExport(string sheetName);

        void SetValue<T>(int rowNumber, int columnNumber, T dataValue);

        void SaveExcelFile();

        byte[] GetExcelFile();
    }
}
