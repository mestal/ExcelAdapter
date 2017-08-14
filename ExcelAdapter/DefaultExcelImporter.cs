using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelProvider
{
    public class DefaultExcelImporter : IExcelImporter
    {
        public IEnumerable<T> Populate<T>(IExcelAdapter excelAdapter, int sheetNumber, Dictionary<string, string> targetSourceList, int startRow, int endRow, Action<T, string, Exception> rowErrorAction) where T : new()
        {
            if (startRow < 1)
                startRow = 1;

            var sheetEndRow = excelAdapter.GetEndRow(sheetNumber);
            if (endRow > sheetEndRow)
                endRow = sheetEndRow;

            IList<T> rows = new List<T>();

            var properties = typeof(T).GetProperties();

            Dictionary<string, int> header = new Dictionary<string, int>();

            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                string castErrorColumn = string.Empty;
                if (rowIndex == 1)
                {
                    header = ExcelHelper.GetExcelHeader(excelAdapter, sheetNumber, rowIndex);
                }
                else
                {
                    var row = new T();
                    try 
                    {
                        foreach (var target in targetSourceList)
                        {
                            try
                            {
                                //key - entity property name
                                //value - excel column
                                var property = properties.First(a => a.Name == target.Key);
                                var parameters = new object[]
                                {
                                    excelAdapter, sheetNumber, header, rowIndex, target.Value
                                };
                                MethodInfo MI = typeof(ExcelHelper).GetMethod("ParseWorksheetValue");
                                MethodInfo genericMethod = MI.MakeGenericMethod(new[] { property.PropertyType });
                                var result = genericMethod.Invoke(null, parameters);

                                property.SetValue(row, result, null);
                            }
                            catch
                            {
                                castErrorColumn = target.Key;
                                throw;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        rowErrorAction(row, castErrorColumn, ex);
                    }
                    rows.Add(row);
                }
            }

            return rows;
        }        
    }
}
