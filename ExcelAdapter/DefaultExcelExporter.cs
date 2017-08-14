using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelProvider
{
    public class DefaultExcelExporter : IExcelExporter
    {
        public void ExportListAsSheet<T>(IExcelAdapter excelAdapter, IEnumerable<T> dataList, string sheetName, Dictionary<string, string> targetSourceList, bool firstRowHeader)
        {
            var properties = typeof(T).GetProperties();
            excelAdapter.AddSheetForExport(sheetName);
            int i = 1;
            int dataStartRow = 2;
            if (firstRowHeader)
            {
                foreach (var targetSource in targetSourceList)
                {
                    excelAdapter.SetValue(1, i, targetSource.Key);
                    i++;
                }
            }
            else
            {
                dataStartRow = 1;
            }

            object dataValue;
            PropertyInfo propInfo;
            foreach (var item in dataList)
            {
                var index = 1;
                foreach (var column in targetSourceList)
                {
                    propInfo = properties.First(a => a.Name == column.Value);
                    dataValue = propInfo.GetValue(item);
                    if (dataValue != null)
                        excelAdapter.SetValue(dataStartRow, index, dataValue.ToString());

                    index++;
                }

                dataStartRow++;
            }

            dataValue = null;
            propInfo = null;

            //sheet.Cells[dataStartRow, 1].LoadFromCollection<T>(list, true);

            excelAdapter.SaveExcelFile();
        }

        public byte[] GetExcelFile(IExcelAdapter excelAdapter)
        {
            return excelAdapter.GetExcelFile();
        }
    }
}
