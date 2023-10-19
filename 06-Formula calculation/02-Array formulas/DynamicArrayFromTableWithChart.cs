using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Entity.Core.Common.CommandTrees.ExpressionBuilder;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlusSamples.FormulaCalculation
{
    public static class DynamicArrayFromTableWithChart
    {
        public static void Run()
        {
            using (var package = new ExcelPackage())
            {
                var sheet1 = package.Workbook.Worksheets.Add("Data");

                // load currency rates from database
                using (var conn = new SQLiteConnection(SampleSettings.ConnectionString))
                {
                    conn.Open();
                    var command = conn.CreateCommand();
                    command.CommandText = "SELECT codeFrom as 'From Currency', codeTo as 'To Currency', date as Date, rate as Rate FROM CurrencyRate";
                    var reader = command.ExecuteReader();
                    var tableRange = sheet1.Cells["A1"].LoadFromDataReader(reader, true, "currencyTable", OfficeOpenXml.Table.TableStyles.Medium1);
                    // set date format for the data in column 3.
                    // we are using the new Skip- and Take-functions that makes access to rows/columns
                    // of a range easier.
                    tableRange
                        .SkipRows(1)
                        .TakeSingleColumn(2)
                        .Style.Numberformat.Format = "yyyy-MM-dd";

                    var sheet2 = package.Workbook.Worksheets.Add("Add dynamic formula");
                    sheet2.Cells["A1"].Formula = "CONCATENATE(\"USD-\",B3)";
                    sheet2.Cells["A1"].Style.Font.Bold = true;
                    // add input field for currency
                    sheet2.Cells["A3"].Value = "Currency";
                    var validation = sheet2.Cells["B3"].DataValidation.AddListDataValidation();
                    validation.Formula.Values.Add("CNY");
                    validation.Formula.Values.Add("DKK");
                    validation.Formula.Values.Add("INR");
                    validation.Formula.Values.Add("EUR");
                    validation.Formula.Values.Add("SEK");
                    sheet2.Cells["B3"].Value = "DKK";
                    sheet2.Cells["B3"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet2.Cells["B3"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet2.Cells["B3"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                    sheet2.Cells["B3"].Style.Font.Bold = true;

                    //Set a dynamic formula to the table headers.
                    sheet2.Cells["A5"].Formula = "Data!A1:D1";
                    sheet2.Cells["A5:D5"].Style.Font.Bold = true;

                    // Here we use the FILTER function to get all USD-DKK rates
                    // from the imported table.                    
                    sheet2.Cells["A6"].Formula = $"FILTER(currencyTable[], currencyTable[To Currency]=B3)";
                    // Dynamic array formulas must always be calculated before saving the workbook.
                    sheet2.Calculate();

                    // The FormulaAddress property contains the range used by the dynamic
                    // array formula after calculation. The variable fa will be used to refer
                    // to address of the dynamic array formulas result range.
                    var fa = sheet2.Cells["A6"].FormulaRange;

                    // set date format for the data in column 3 of the dynamic array.
                    // we are using the new TakeSingleColumn function which provides easier
                    // access to entire columns.
                    fa.TakeSingleColumn(2).Style.Numberformat.Format = "yyyy-MM-dd";
                    // Now let's add a chart for the filtered array (initially showing USD-DKK rates)
                    var chart = sheet2.Drawings.AddLineChart("Dynamic Chart", eLineChartType.Line);
                    chart.Title.LinkedCell = sheet2.Cells["B3"];
                    var series = chart.Series.Add(
                        fa.TakeSingleColumn(3),
                        fa.TakeSingleColumn(2)
                    );
                    series.Header = "Rate";

                    //Add conditional formatting for each currency in the filtered data.
                    AddConditionalNumberFormat(sheet2.Cells["D5:D1000"], "$B5=\"CNY\"", "[$¥-804]#,##0.00");
                    AddConditionalNumberFormat(sheet2.Cells["D5:D1000"], "$B5=\"DKK\"", "#,##0.00\\ [$kr.-406]");
                    AddConditionalNumberFormat(sheet2.Cells["D5:D1000"], "$B5=\"EUR\"", "#,##0.00\\ [$€-1]");
                    AddConditionalNumberFormat(sheet2.Cells["D5:D1000"], "$B5=\"INR\"", "[$₹-4009]\\ #,##0.00");
                    AddConditionalNumberFormat(sheet2.Cells["D5:D1000"], "$B5=\"SEK\"", "#,##0.00\\ [$kr-41D]");

                    chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle7);

                    chart.SetPosition(1, 0, 6, 0);
                    chart.SetSize(200);
                    sheet1.Cells.AutoFitColumns();
                    sheet2.Cells.AutoFitColumns();
                    
                    sheet2.Select("B3", true);
                }
                package.SaveAs(FileUtil.GetCleanFileInfo("6.2-DynamicArrayFormulasWithChart.xlsx"));
            }
                
        }

        private static void AddConditionalNumberFormat(ExcelRangeBase range, string formula, string numberFormat)
        {
            var cf1 = range.ConditionalFormatting.AddExpression();
            cf1.Formula = formula;
            cf1.Style.NumberFormat.Format = numberFormat;
        }
    }
}
