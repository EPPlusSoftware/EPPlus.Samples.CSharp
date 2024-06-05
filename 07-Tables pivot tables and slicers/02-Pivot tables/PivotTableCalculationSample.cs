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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Data.SqlClient;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Data.SQLite;
using EPPlusSamples.FiltersAndValidations;
using OfficeOpenXml.Table.PivotTable.Calculation;
namespace EPPlusSamples.PivotTables
{
    /// <summary>
    /// This class shows how to calculate pivottables and fetch data via the CalculatedData propety or the GetPivotData method.
    /// </summary>
    public static class PivotTablesCalculationSample 
    {
        public static void Run()
        {
            Console.WriteLine("Running sample 7.3-Pivot Table Calculation");

            FileInfo templateFile = FileUtil.GetFileInfo("7.2-PivotTables.xlsx");
            if(templateFile.Exists==false)
            {
                Console.WriteLine("Template file 7.2-PivotTables.xlsx does not exist. Please make sure the sample PivotTablesSample.Run() sample has executed to create this file.");
            }
            using (ExcelPackage pck = new ExcelPackage(templateFile))
            {
                var pt1 = pck.Workbook.Worksheets["PivotSimple"].PivotTables[0];
                var pt2 = pck.Workbook.Worksheets["PivotDateGrp"].PivotTables[0];
                var pt3 = pck.Workbook.Worksheets["PivotWithPageField"].PivotTables[0];
                var pt4 = pck.Workbook.Worksheets["PivotWithSlicer"].PivotTables[0];
                var pt5 = pck.Workbook.Worksheets["PivotWithCalculatedField"].PivotTables[0];
                var pt6 = pck.Workbook.Worksheets["PivotWithCaptionFilter"].PivotTables[0];
                var pt7 = pck.Workbook.Worksheets["PivotWithShowAsFields"].PivotTables[0];
                var pt8 = pck.Workbook.Worksheets["PivotSorting"];

                //Use calculate on the pivot table to calculate the values that can be accessed via the CalculatedData property or the GetPivotData method.
                //If no calculation has been performed, EPPlus will call this method when using these properties, but will not refresh the pivot cache unless it does not exist.
                //If you have altered data in the pivot source, make sure to call this method with true, to update the pivot cache.
                //Also make sure to calculate any formulas that the pivot table source contains before calculating the data.
                StandardPivotTableSample(pt1);
                DateGroupSample(pt2);
                PageFieldFilterSample(pt3);
                SlicerFilterSample(pt4);
                CalculatedFieldSample(pt5);
                CaptionFilterSample(pt6);
                ShowAsSample(pt7);
                GetPivotDataMethodSample(pck.Workbook.Worksheets["PivotSorting"]);
            }
        }

        private static void StandardPivotTableSample(ExcelPivotTable pt)
        {
            pt.Calculate(true);

            var tot = pt.CalculatedData.GetValue(); //Get the grant total from the pivot table.
            var capVerde = pt.CalculatedData.SelectField("Country", "Cape verde").GetValue();

            Console.WriteLine($"The calculated grand total for pivot table {pt.Name} is: {tot:N0}");
            Console.WriteLine($"The total for Cap Verde for pivot table {pt.Name} is: {capVerde:N0}");
            Console.WriteLine();
        }

        private static void DateGroupSample(ExcelPivotTable pt)
        {
            var hellenKuhlman = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").
                                GetValue("OrderValue");

            var hellenKuhlman2017Q3Tax = pt.CalculatedData.
                               SelectField("Name", "Hellen Kuhlman").
                               SelectField("Years", "2017").
                               SelectField("Quarters", "Q3").
                                GetValue("Tax");

            Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt.Name} is: {hellenKuhlman:N0}");
            Console.WriteLine($"The value for Tax for Hellen Kuhlman,Q3 2017 for pivot table {pt.Name} is: {hellenKuhlman2017Q3Tax:N2}");
            Console.WriteLine();
        }

        private static void PageFieldFilterSample(ExcelPivotTable pt)
        {
            object hellenKuhlman = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").
                                GetValue("OrderValue");
            var hellenKuhlman2017Q4Tax = pt.CalculatedData.
                               SelectField("Name", "Hellen Kuhlman").
                               SelectField("OrderDate", "Qtr4").
                                GetValue("Tax");

            Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt.Name} is: {hellenKuhlman:N0}. This value has been filtered by the page field.");
            Console.WriteLine($"The value for Tax for Hellen Kuhlman,Q4 2017 for pivot table {pt.Name} is: {hellenKuhlman2017Q4Tax:N2}");
            Console.WriteLine();
        }

        private static void SlicerFilterSample(ExcelPivotTable pt)
        {
            object hellenKuhlman = pt.CalculatedData.SelectField("Name", "Hellen Kuhlman").
                                GetValue("OrderValue");
            var walschSum = pt.CalculatedData.SelectField("Name", "Walsh LLC").
                                GetValue("OrderValue");

            Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt.Name} is: {hellenKuhlman:N0}. This value has been filtered by the page field.");
            Console.WriteLine($"The value for OrderValue for Walsh LLC,Q4 2017 for pivot table {pt.Name} is: {walschSum:N2}. It is filtered out by the slicer.");
            Console.WriteLine();
        }

        private static void CalculatedFieldSample(ExcelPivotTable pt)
        {
            var sengerOrderValue = pt.CalculatedData.
                                SelectField("CompanyName", "Senger LLC").
                                GetValue("OrderValue");

            var sengerTax = pt.CalculatedData.
                                SelectField("CompanyName", "Senger LLC").
                                GetValue("Tax");

            var sengerFreight = pt.CalculatedData.
                                SelectField("CompanyName", "Senger LLC").
                                GetValue("Freight");

            var sengerTotal = pt.CalculatedData.
                                SelectField("CompanyName", "Senger LLC").
                                GetValue("Total");

            var grandTotal = pt.CalculatedData.
                                GetValue("Total");

            Console.WriteLine($"Calculated Fields: The value of field OrderValue for Senger LLC for pivot table {pt.Name} is: {sengerOrderValue:N0}.");
            Console.WriteLine($"Calculated Fields: The value of field Tax for Senger LLC for pivot table {pt.Name} is: {sengerTax:N0}.");
            Console.WriteLine($"Calculated Fields: The value of field Freight for Senger LLC for pivot table {pt.Name} is: {sengerFreight:N0}.");
            Console.WriteLine($"Calculated Fields: The value of field Total for Senger LLC for pivot table {pt.Name} is: {sengerTotal:N0}. This field uses the formula [OrderValue]+[Tax]+[Freight]");
            
            Console.WriteLine($"Calculated Fields: The grand value for OrderValue for  2017 for pivot table {pt.Name} is: {grandTotal:N2}.");
            Console.WriteLine();
        }
        private static void CaptionFilterSample(ExcelPivotTable pt)
        {
            //Sabryna Schulist
            var sabrynaSchulistOrderValue = pt.CalculatedData.
                                SelectField("Name", "Sabryna Schulist").
                                GetValue("OrderValue");

            var orderDate = new DateTime(2017, 8, 27, 1, 57, 0);

            var sabrynaSchulistDateTimeTax = pt.CalculatedData.
                                SelectField("Name", "Sabryna Schulist").
                                SelectField("OrderDate", orderDate).
                                GetValue("Tax");

            //Chelsey Powlowski - is filtered out by the caption filter as the name startes with "C". #REF! will be returened
            var chelseyPowlowskiOrderValue = pt.CalculatedData.
                                SelectField("Name", "Chelsey Powlowski").
                                GetValue("OrderValue");

            //Get the grand total for field OrderValue
            var grandTotalOrderValue = pt.CalculatedData.
                                GetValue("OrderValue");

            Console.WriteLine($"Caption Filters: The value of field OrderValue for Sabryna Schulist for pivot table {pt.Name} is: {sabrynaSchulistOrderValue:N0}.");
            Console.WriteLine($"Caption Filters: The value of field Tax Sabryna Schulist, {orderDate} for pivot table {pt.Name} is: {sabrynaSchulistDateTimeTax:N0}.");
            
            Console.WriteLine($"Caption Filters: The value of field OrderValue for Chelsey Powlowski for pivot table {pt.Name} is: {chelseyPowlowskiOrderValue:N0}. This value has been filtered out by the caption filter");
            Console.WriteLine($"Caption Filters: The grand total of field OrderValue for pivot table {pt.Name} is: {grandTotalOrderValue:N0}. All Name's starting with \"C\" has been filtered out by the caption filter");
            Console.WriteLine();
        }
        private static void ShowAsSample(ExcelPivotTable pt)
        {
            var wizaHauckEUR = pt.CalculatedData.
                                SelectField("CompanyName", "Wiza-Hauck").
                                SelectField("Currency", "EUR").
                                GetValue("Order value");

            var wizaHauckEURPercentOfTotal = pt.CalculatedData.
                                SelectField("CompanyName", "Wiza-Hauck").
                                //SelectField("Name", "Kianna Bradtke").
                                SelectField("Currency", "EUR").
                                GetValue("Order value % of total");

            var wizaHauckEURCountDifferance = pt.CalculatedData.
                                SelectField("CompanyName", "Wiza-Hauck").
                                SelectField("Currency", "EUR").
                                GetValue("Count Difference From Previous");

            Console.WriteLine($"Show as: The value of field Order value for Wiza-Hauck, EUR for pivot table {pt.Name} is: {wizaHauckEUR:N0}.");
            Console.WriteLine($"Show as: The value of field Order value % of total for Wiza-Hauck, EUR for pivot table {pt.Name} is: {wizaHauckEURPercentOfTotal:P1}.");
            Console.WriteLine($"Show as: The value of field Order value From Previous for Wiza-Hauck, EUR for pivot table {pt.Name} is: {wizaHauckEURCountDifferance:N0}.");
            Console.WriteLine();
        }
        /// <summary>
        /// This sample shows how to use the ExcelPivotTable.GetPivotData method as an option to use the ExcelPivotTable.CalculatedData property.
        /// </summary>
        /// <param name="ws">The worksheet containing the pivot tables</param>
        private static void GetPivotDataMethodSample(ExcelWorksheet ws)
        {
            var pt1 = ws.PivotTables[0];
            var pt2 = ws.PivotTables[1];
            var pt3 = ws.PivotTables[2];

            var grandTotal = pt1.GetPivotData("OrderValue");
            var tajikistanTotal = pt2.GetPivotData("OrderValue", new List<PivotDataFieldItemSelection>() { new PivotDataFieldItemSelection("Country", "Tajikistan")  });
            var equatorialGuinea = pt3.GetPivotData("OrderValue", new List<PivotDataFieldItemSelection>() { new PivotDataFieldItemSelection("Country", "Equatorial Guinea") });

            var equatorialGuineaChelseyPowlowski = pt3.GetPivotData("OrderValue", new List<PivotDataFieldItemSelection>() 
                { 
                new PivotDataFieldItemSelection("Country", "Equatorial Guinea"),
                new PivotDataFieldItemSelection("Name", "Chelsey Powlowski")
                });

            Console.WriteLine($"GetPivotData method: The grand total for pivot table {pt1.Name} is: {grandTotal:N0}.");
            Console.WriteLine($"GetPivotData method: The value of field OrderValue for Tajikistan for pivot table {pt2.Name} is: {tajikistanTotal:N0}.");
            Console.WriteLine($"GetPivotData method: The value of field OrderValue for Equatorial Guinea, Chelsey Powlowski for pivot table {pt3.Name} is: {equatorialGuineaChelseyPowlowski:N0}.");
            Console.WriteLine();
        }
    }
}