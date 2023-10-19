using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Net;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class EqualBetweenExpressionsExamples
    {
        public static void Run(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("ExpressionExamples");

            var range = "B1:B30";

            sheet.Cells["B1:B2"].Value = 3;
            sheet.Cells["B3"].Value = 6;
            sheet.Cells["B4"].Value = 35;
            sheet.Cells["B5"].Value = 36;
            sheet.Cells["B6"].Value = 37;
            sheet.Cells["B7"].Value = 38;
            sheet.Cells["B8"].Value = 68;
            sheet.Cells["B8"].Value = 444;
            sheet.Cells["B9"].Value = 1000;
            sheet.Cells["B10"].Value = 25;
            sheet.Cells["B11"].Value = 43;
            sheet.Cells["B12"].Value = 43;
            sheet.Cells["B13"].Value = 43;

            // -------------------------------------------------------------------
            // Create a Between rule
            // -------------------------------------------------------------------
            var betweenRule = sheet.ConditionalFormatting.AddBetween(
              range);

            betweenRule.Formula = "IF(A1>5,10,20)";
            betweenRule.Formula2 = "IF(A1>5,30,50)";

            betweenRule.Style.Border.Right.Style = ExcelBorderStyle.Thick;
            betweenRule.Style.Border.Right.Color.Color = Color.Goldenrod;

            // -------------------------------------------------------------------
            // Create an Equal rule
            // -------------------------------------------------------------------
            var equal = sheet.ConditionalFormatting.AddEqual(
              range);

            equal.Formula = "6";

            equal.Style.Border.Left.Style = ExcelBorderStyle.MediumDashed;
            equal.Style.Border.Left.Color.Color = Color.Purple;

            // -------------------------------------------------------------------
            // Create an NotEqual rule
            // -------------------------------------------------------------------
            var notEqual = sheet.ConditionalFormatting.AddNotEqual(
              "A10:A11");

            notEqual.Formula = "14";

            notEqual.Style.Border.BorderAround(ExcelBorderStyle.DashDotDot, Color.Firebrick);

            sheet.Cells["A10"].Value = 14;
            sheet.Cells["A11"].Value = 10;

            // -------------------------------------------------------------------
            // Create an Expression rule
            // -------------------------------------------------------------------
            var customExpression = sheet.ConditionalFormatting.AddExpression(
              range);

            customExpression.Formula = "B1=B2";
            customExpression.Style.Font.Bold = true;

            // -------------------------------------------------------------------
            // Create a GreaterThan rule
            // -------------------------------------------------------------------
            var greater = sheet.ConditionalFormatting.AddGreaterThan(
              range);

            greater.Formula = "SE(B1<10,10,65)";

            greater.Style.Fill.PatternType = ExcelFillStyle.Solid;
            greater.Style.Fill.BackgroundColor.Color = Color.DarkOrchid;

            // -------------------------------------------------------------------
            // Create a GreaterThanOrEqual rule
            // -------------------------------------------------------------------
            var greaterEqual = sheet.ConditionalFormatting.AddGreaterThanOrEqual(
              range);

            greaterEqual.Formula = "40";

            greaterEqual.Priority = 1;
            greaterEqual.Style.Border.BorderAround(ExcelBorderStyle.Double, Color.Red);

            // -------------------------------------------------------------------
            // Create a LessThan rule
            // -------------------------------------------------------------------
            var lessThan = sheet.ConditionalFormatting.AddLessThan(
              range);

            lessThan.Formula = "36";
            lessThan.Style.Font.Strike = true;

            // -------------------------------------------------------------------
            // Create a LessThanOrEqual rule
            // -------------------------------------------------------------------
            var lessThanEqual = sheet.ConditionalFormatting.AddLessThanOrEqual(
              range);

            lessThanEqual.Formula = "37";
            lessThanEqual.Style.Font.Italic = true;

            // -------------------------------------------------------------------
            // Create a NotBetween rule
            // -------------------------------------------------------------------
            var notBetween = sheet.ConditionalFormatting.AddNotBetween(
              range);

            notBetween.Style.Font.Color.Color = Color.ForestGreen;

            notBetween.Formula = "333";
            notBetween.Formula2 = "999";
        }
    }
}
