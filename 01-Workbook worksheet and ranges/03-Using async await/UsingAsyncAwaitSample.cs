/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using System;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Threading;
using OfficeOpenXml.Table;
using System.Text;
using EPPlusSamples.FiltersAndValidations;

namespace EPPlusSamples.WorkbookWorksheetAndRanges
{
    class UsingAsyncAwaitSample
    {
        /// <summary>
        /// Shows a few different ways to load / save asynchronous
        /// </summary>
        public static async Task RunAsync()
        {
            Console.WriteLine("Running sample 1.3-Async-Await");
            var file = FileUtil.GetCleanFileInfo("1.03-AsyncAwait.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");

                using (var sqlConn = new SQLiteConnection(SampleSettings.ConnectionString))
                {
                    sqlConn.Open();
                    var sql = SqlStatements.OrdersSql;
                    using (var sqlCmd = new SQLiteCommand(sql, sqlConn))
                    {
                        var range = await ws.Cells["B2"].LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), true, "Table1", TableStyles.Medium10);
                        range.AutoFitColumns();
                    }
                }

                await package.SaveAsync();
            }

            //Load the package async again.
            using (var package = new ExcelPackage())
            {
                await package.LoadAsync(file);

                var newWs = package.Workbook.Worksheets.Add("AddedSheet2");
                var range = await newWs.Cells["A1"].LoadFromTextAsync(FileUtil.GetFileInfo("01-Workbook Worksheet and Ranges\\03-Using Async Await", "Importfile.txt"), new ExcelTextFormat { Delimiter='\t' });
                range.AutoFitColumns();

                await package.SaveAsAsync(FileUtil.GetCleanFileInfo("1.03-AsyncAwait-LoadedAndModified.xlsx"));
            }
            Console.WriteLine("Sample 1.3 created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();

        }
    }
}
