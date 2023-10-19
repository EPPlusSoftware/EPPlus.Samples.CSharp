using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class StandardDeviationTopDown
    {
        public static void Run(ExcelPackage pck)
        {
            var worksheet = pck.Workbook.Worksheets.Add("StdDev_TopBottom");

            worksheet.Cells["B1:B43"].Formula = "ROW()";
                
            // -------------------------------------------------------------------
            // Create a Above StdDev rule
            // -------------------------------------------------------------------
            var zeroDeviation = worksheet.ConditionalFormatting.AddAboveStdDev(
                new ExcelAddress("B1:B43"));
            zeroDeviation.StdDev = 0;

            zeroDeviation.Style.Font.Bold = true;

            // -------------------------------------------------------------------
            // Create a Below StdDev rule
            // -------------------------------------------------------------------
            var twoDeviation = worksheet.ConditionalFormatting.AddBelowStdDev(
                new ExcelAddress("B1:B43"));

            twoDeviation.StdDev = 2;
            twoDeviation.Style.Fill.PatternType = ExcelFillStyle.Solid;
            twoDeviation.Style.Fill.BackgroundColor.Color = Color.ForestGreen;

            //Make a single cell actually exist at stdev 2
            worksheet.Cells["B14"].Value = -177;

            // -------------------------------------------------------------------
            // Create a Bottom rule
            // -------------------------------------------------------------------
            var bottomRank4 = worksheet.ConditionalFormatting.AddBottom(
                new ExcelAddress("B1:B43"));

            bottomRank4.Rank = 4;

            bottomRank4.Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.MediumVioletRed);

            // -------------------------------------------------------------------
            // Create a Bottom Percent rule
            // -------------------------------------------------------------------
            var bottomPercent15 = worksheet.ConditionalFormatting.AddBottomPercent(
                new ExcelAddress("B1:B43"));

            bottomPercent15.Rank = 15;

            bottomPercent15.Style.Fill.PatternType = ExcelFillStyle.Solid;
            bottomPercent15.Style.Fill.BackgroundColor.Color = Color.DeepSkyBlue;

            // -------------------------------------------------------------------
            // Create a Top rule
            // -------------------------------------------------------------------
            var top = worksheet.ConditionalFormatting.AddTop(
                new ExcelAddress("B1:B43"));
            top.Style.Fill.PatternType = ExcelFillStyle.Solid;
            top.Style.Fill.BackgroundColor.Color = Color.MediumPurple;

            // -------------------------------------------------------------------
            // Create a Top Percent rule
            // -------------------------------------------------------------------
            var topPercent = worksheet.ConditionalFormatting.AddTopPercent(
                new ExcelAddress("B1:B43"));

            topPercent.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            topPercent.Style.Border.Left.Color.Theme = eThemeSchemeColor.Text2;
            topPercent.Style.Border.Bottom.Style = ExcelBorderStyle.DashDot;
            topPercent.Style.Border.Bottom.Color.SetColor(ExcelIndexedColor.Indexed8);
            topPercent.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            topPercent.Style.Border.Right.Color.Color = Color.Blue;
            topPercent.Style.Border.Top.Style = ExcelBorderStyle.Hair;
            topPercent.Style.Border.Top.Color.Auto = true;

            worksheet.Cells.AutoFitColumns();
        }
    }
}
