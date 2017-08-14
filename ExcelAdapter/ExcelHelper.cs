using System.Collections.Generic;
using System.ComponentModel;

namespace ExcelProvider
{
    public class ExcelHelper
    {
        public static Dictionary<string, int> GetExcelHeader(IExcelAdapter excelAdapter, int sheetNumber, int rowIndex)
        {
            Dictionary<string, int> header = new Dictionary<string, int>();

            for (int columnIndex = excelAdapter.GetStartColumn(sheetNumber); columnIndex <= excelAdapter.GetEndColumn(sheetNumber); columnIndex++)
            {
                if (excelAdapter.GetCellValue(sheetNumber, columnIndex, rowIndex) != null)
                {
                    string columnName = excelAdapter.GetCellValue(sheetNumber, columnIndex, rowIndex).ToString();

                    if (!header.ContainsKey(columnName) && !string.IsNullOrEmpty(columnName))
                    {
                        header.Add(columnName, columnIndex);
                    }
                }

            }

            return header;
        }

        public static T ParseWorksheetValue<T>(IExcelAdapter excelAdapter, int sheetNumber, Dictionary<string, int> header, int rowIndex, string columnName)
        {
            string value = string.Empty;
            int? columnIndex = header.ContainsKey(columnName) ? header[columnName] : (int?)null;

            if (columnIndex != null && excelAdapter.GetCellValue(sheetNumber, columnIndex.Value, rowIndex) != null)
            {
                value = excelAdapter.GetCellValue(sheetNumber, columnIndex.Value, rowIndex).ToString();
            }

            var converter = TypeDescriptor.GetConverter(typeof(T));
            var result = converter.ConvertFrom(value);

            return (T)result;
        }
    }
}
