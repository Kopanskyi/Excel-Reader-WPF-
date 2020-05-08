using System;
using System.Data;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace ExcelReader
{
    public static class ExcelFileReader
    {
        public static DataTable ReadFile(string filePath)
        {
            var xlApp = new Application();
            var xlWorkbooks = xlApp.Workbooks;
            var xlWorkbook = xlWorkbooks.Open(filePath);

            var xlWorksheet = xlWorkbook.Sheets[1];
            var xlRange = xlWorksheet.UsedRange;

            DataTable dataTable = CreateDataTable(xlWorksheet);
            
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);

            return dataTable;
        }

        private static DataTable CreateDataTable(_Worksheet xlWorksheet)
        {
            var dataTable = new DataTable();
            var worksheetName = xlWorksheet.Name;
            dataTable.TableName = worksheetName;
            var xlRange = xlWorksheet.UsedRange;
            var valueArray = (object[,])xlRange.Value[XlRangeValueDataType.xlRangeValueDefault];

            for (var k = 1; k <= valueArray.GetLength(1); k++)
            {
                dataTable.Columns.Add(valueArray[1, k].ToString());
            }

            var singleDValue = new object[valueArray.GetLength(1)];

            for (var i = 2; i <= valueArray.GetLength(0); i++)
            {
                for (var j = 0; j < valueArray.GetLength(1); j++)
                {
                    singleDValue[j] = valueArray[i, j + 1];
                }

                dataTable.LoadDataRow(singleDValue, LoadOption.PreserveChanges);
            }

            return dataTable;
        }
    }
}
