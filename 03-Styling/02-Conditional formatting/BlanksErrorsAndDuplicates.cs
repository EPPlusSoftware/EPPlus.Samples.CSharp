using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class BlanksErrorsAndDuplicates
    {
        public static void Run(ExcelPackage pck)
        {
            var sheet = pck.Workbook.Worksheets.Add("BlanksAndErrors");

            var address = "A1:A20";

            // -------------------------------------------------------------------
            // Create a ContainsBlanks rule
            // -------------------------------------------------------------------
            var containsBlanks = sheet.ConditionalFormatting.AddContainsBlanks(
              address);

            containsBlanks.Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.Goldenrod);

            // -------------------------------------------------------------------
            // Create a NotContainsBlanks rule
            // -------------------------------------------------------------------
            var noBlanks = sheet.ConditionalFormatting.AddNotContainsBlanks(
              address);

            noBlanks.Style.Border.Top.Style = ExcelBorderStyle.Double;
            noBlanks.Style.Border.Top.Color.Color = Color.ForestGreen;

            sheet.Cells["A3:A6"].Formula = "Row()";

            // -------------------------------------------------------------------
            // Create a ContainsErrors rule
            // -------------------------------------------------------------------
            var containsErrors = sheet.ConditionalFormatting.AddContainsErrors(
              address);

            //Add a few incorrect formulas
            sheet.Cells["A7"].Formula = "I an Invalid Formula";
            sheet.Cells["A8"].Formula = "SUM(1,\"Nonsense\")";
            //Will show up appropriately but prompts excel to update links on opening the file
            //sheet.Cells["A9"].Formula = "SUM(1,nonExistent!J12)";

            containsErrors.Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Red);

            containsErrors.Priority = 1;

            // -------------------------------------------------------------------
            // Create a NotContainsErrors rule
            // -------------------------------------------------------------------
            var noErrors = sheet.ConditionalFormatting.AddNotContainsErrors(
              address);

            noErrors.Style.Border.Right.Style = ExcelBorderStyle.Double;
            noErrors.Style.Border.Right.Color.Color = Color.Purple;

            // -------------------------------------------------------------------
            // Create a DuplicateValues rule
            // -------------------------------------------------------------------
            var duplicateValues = sheet.ConditionalFormatting.AddDuplicateValues(
              address);

            duplicateValues.Style.Fill.Style = eDxfFillStyle.PatternFill;
            duplicateValues.Style.Fill.PatternType = ExcelFillStyle.Solid;
            duplicateValues.Style.Fill.BackgroundColor.Color = Color.DarkOrange;

        }
    }
}
