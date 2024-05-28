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

namespace EPPlusSamples.PivotTables
{
    /// <summary>
    /// This class shows how to use pivottables 
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
                //var pt2 = CreatePivotTableWithDataGrouping(pck, dataRange);
                //var pt3 = CreatePivotTableWithPageFilter(pck, pt2.CacheDefinition);
                //var pt4 = CreatePivotTableWithASlicer(pck, pt2.CacheDefinition);
                //var pt5 = CreatePivotTableWithACalculatedField(pck, pt2.CacheDefinition);
                //var pt6 = CreatePivotTableCaptionFilter(pck, dataRange);
                //var pt7 = CreatePivotTableWithDataFieldsUsingShowAs(pck, dataRange);


                //Use calculate on the pivot table to calculate the values that can be accessed via the CalculatedData property or the GetPivotData method.
                //If no calculation has been performed, EPPlus will call this method when using these properties, but will not refresh the pivot cache unless it does not exist.
                //If you have altered data in the pivot source, make sure to call this method with true, to update the pivot cache.
                //Also make sure to calculate any formulas that the pivot table source contains before calculating the data.
                pt1.Calculate(true); 

                var tot = pt1.CalculatedData.GetValue(); //Get the grant total from the pivot table.
                var capVerde = pt1.CalculatedData.SelectField("Country", "Cape verde").GetValue();

                Console.WriteLine($"The calculated grand total for pivot table {pt1.Name} is: {tot:N0}");
                Console.WriteLine($"The total for Cap Verde for pivot table {pt1.Name} is: {capVerde:N0}");


                var hellenKuhlman = pt2.CalculatedData.SelectField("Name", "Hellen Kuhlman").
                                    GetValue("OrderValue");

                var hellenKuhlman2017Q3Tax = pt2.CalculatedData.SelectField("Name", "Hellen Kuhlman").
                                   SelectField("Years","2017").
                                   SelectField("Quarters","Q3").
                                    GetValue("Tax");

                Console.WriteLine($"The Total for OrderValue for Hellen Kuhlman for pivot table {pt2.Name} is: {hellenKuhlman:N0}");
                Console.WriteLine($"The value for Tax for Hellen Kuhlman,Q3 2017 for pivot table {pt2.Name} is: {hellenKuhlman2017Q3Tax:N2}");

            }
        }
    }
}