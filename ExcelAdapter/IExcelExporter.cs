using System.Collections.Generic;

namespace ExcelProvider
{
    public interface IExcelExporter
    {
        void ExportListAsSheet<T>(IExcelAdapter excelAdapter, IEnumerable<T> dataList, string sheetName,
            Dictionary<string, string> targetSourceList, bool firstRowHeader);

        byte[] GetExcelFile(IExcelAdapter excelAdapter);
    }
}
