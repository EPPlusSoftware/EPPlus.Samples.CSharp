using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class AveragesExample
    {
        public static void Run(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("AverageExamples");

            sheet.Cells["A1:B21"].Formula = "ROW()";

            // -------------------------------------------------------------------
            // Create an Above Average rule
            // -------------------------------------------------------------------
            var above = sheet.ConditionalFormatting.AddAboveAverage(
                new ExcelAddress("A1:B21"));

            //Properties allow you to change the formatting of conditional formattings
            //Multiple font properties can be changed to alter the apperance of the formatting
            above.Style.Font.Bold = true;
            above.Style.Font.Color.Color = Color.Red;
            above.Style.Font.Strike = true;

            // -------------------------------------------------------------------
            // Create an Above Or Equal Average rule
            // -------------------------------------------------------------------
            var aboveOrEqual = sheet.ConditionalFormatting.AddAboveOrEqualAverage(
                new ExcelAddress("A1:A21"));

            //Other properties like style can change background color
            aboveOrEqual.Style.Fill.PatternType = ExcelFillStyle.Solid;
            aboveOrEqual.Style.Fill.BackgroundColor.Color = Color.DarkBlue;

            // -------------------------------------------------------------------
            // Create a Below Average rule
            // -------------------------------------------------------------------
            var belowAverage = sheet.ConditionalFormatting.AddBelowAverage(
                new ExcelAddress("A1:B21"));

            belowAverage.Style.Fill.PatternType = ExcelFillStyle.Solid;
            belowAverage.Style.Fill.BackgroundColor.Color = Color.DarkRed;

            // -------------------------------------------------------------------
            // Create a Below Or Equal Average rule
            // -------------------------------------------------------------------
            var belowOrEqual = sheet.ConditionalFormatting.AddBelowOrEqualAverage(
                new ExcelAddress("A1:B21"));
            belowOrEqual.Style.Font.Color.Color = Color.White;

            belowOrEqual.Style.Fill.PatternType = ExcelFillStyle.Solid;
            belowOrEqual.Style.Fill.BackgroundColor.Color = Color.DarkGreen;

            //Note that when two properties conflict like belowEqual and aboveEqual on the background color the one with the lowest priority number "wins"
            //Test switching them around and watch the A11 cells closely.
            belowOrEqual.Priority = 2;
            aboveOrEqual.Priority = 1;

            sheet.Cells.AutoFitColumns();
        }
    }
}
