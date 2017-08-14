using System;
using System.Collections.Generic;

namespace ExcelProvider
{
    public interface IExcelImporter
    {
        IEnumerable<T> Populate<T>(IExcelAdapter excelAdapter, int sheetNumber, Dictionary<string, string> targetSourceList, int startRow, int endRow, Action<T, string, Exception> rowErrorAction) where T : new();
    }
}
