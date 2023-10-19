using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class DatesAndTime
    {
        public static void Run(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("DatesAndTimeExamples");

            AddDatesAndTimesToSheet(sheet);

            // -------------------------------------------------------------------
            // Create a Last 7 Days rule
            // -------------------------------------------------------------------
            //ExcelAddress timePeriodAddress = new ExcelAddress("B21:E34 A11:A20");
            ExcelAddress timePeriodAddress = new ExcelAddress("A1:K40");

            var last7Days = sheet.ConditionalFormatting.AddLast7Days("A1:A40");

            last7Days.Style.Fill.PatternType = ExcelFillStyle.LightTrellis;
            last7Days.Style.Fill.PatternColor.Color = Color.BurlyWood;
            last7Days.Style.Fill.BackgroundColor.Color = Color.LightCyan;

            // -------------------------------------------------------------------
            // Create a Last Week rule
            // -------------------------------------------------------------------
            var lastWeek = sheet.ConditionalFormatting.AddLastWeek("B1:B40");
            //lastWeek.Style.NumberFormat.Format = "YYYY";
            lastWeek.Style.Fill.PatternType = ExcelFillStyle.Solid;
            lastWeek.Style.Fill.BackgroundColor.Color = Color.Orange;

            // -------------------------------------------------------------------
            // Create a This Week rule
            // -------------------------------------------------------------------
            var thisWeek = sheet.ConditionalFormatting.AddThisWeek("B1:B40");

            thisWeek.Style.Fill.PatternType = ExcelFillStyle.Solid;
            thisWeek.Style.Fill.BackgroundColor.Color = Color.YellowGreen;

            // -------------------------------------------------------------------
            // Create a Next Week rule
            // -------------------------------------------------------------------
            var nextWeek = sheet.ConditionalFormatting.AddNextWeek("B1:B40");

            nextWeek.Style.Fill.PatternType = ExcelFillStyle.Solid;
            nextWeek.Style.Fill.BackgroundColor.Color = Color.ForestGreen;

            // -------------------------------------------------------------------
            // Create a Today rule
            // -------------------------------------------------------------------
            var today = sheet.ConditionalFormatting.AddToday("C1:C40");

            today.Style.Fill.PatternType = ExcelFillStyle.Solid;
            today.Style.Fill.BackgroundColor.Color = Color.Gold;

            today.Style.Border.BorderAround(ExcelBorderStyle.Thick, Color.Gold);

            // -------------------------------------------------------------------
            // Create a Tomorrow rule
            // -------------------------------------------------------------------
            var tomorrow = sheet.ConditionalFormatting.AddTomorrow("C1:C40");

            tomorrow.Style.Fill.PatternType = ExcelFillStyle.Solid;
            tomorrow.Style.Fill.BackgroundColor.Color = Color.LightSkyBlue;

            tomorrow.Style.Border.BorderAround(ExcelBorderStyle.DashDot, Color.Violet);

            // -------------------------------------------------------------------
            // Create a Yesterday rule
            // -------------------------------------------------------------------
            var yesterday = sheet.ConditionalFormatting.AddYesterday("C1:C40");

            yesterday.Style.Fill.PatternType = ExcelFillStyle.Solid;
            yesterday.Style.Fill.BackgroundColor.Color = Color.DimGray;

            yesterday.Style.Border.BorderAround(ExcelBorderStyle.Dashed, Color.DarkRed);

            // -------------------------------------------------------------------
            // Create a Last Month rule
            // -------------------------------------------------------------------
            var lastMonth = sheet.ConditionalFormatting.AddLastMonth("E1:E40");

            //lastMonth.Style.NumberFormat.Format = "YYYY";
            lastMonth.Style.Fill.PatternType = ExcelFillStyle.Solid;
            lastMonth.Style.Fill.BackgroundColor.Color = Color.OrangeRed;

            // -------------------------------------------------------------------
            // Create a This Month rule
            // -------------------------------------------------------------------
            var thisMonth = sheet.ConditionalFormatting.AddThisMonth("F1:F40");

            thisMonth.Style.Fill.PatternType = ExcelFillStyle.Solid;
            thisMonth.Style.Fill.BackgroundColor.Color = Color.LightGoldenrodYellow;

            // -------------------------------------------------------------------
            // Create a Next Month rule
            // -------------------------------------------------------------------
            var nextMonth = sheet.ConditionalFormatting.AddNextMonth("G1:G40");

            nextMonth.Style.Fill.PatternType = ExcelFillStyle.Solid;
            nextMonth.Style.Fill.BackgroundColor.Color = Color.ForestGreen;

            sheet.Cells.AutoFitColumns();
        }

        private static void AddDatesAndTimesToSheet(ExcelWorksheet sheet)
        {
            int startOffset = (int)DateTime.Now.DayOfWeek;
            var lastWeekDate = DateTime.Now.AddDays(-7 - startOffset);
            var year = $"{DateTime.Now.Year}";

            string lastMonth = $"{year}-{DateTime.Now.AddMonths(-1).Month}-";
            string thisMonth = $"{year}-{DateTime.Now.Month}-";
            string nextMonth = $"{year}-{DateTime.Now.AddMonths(+1).Month}-";

            for (int i = 1; i < 11; i++)
            {
                sheet.Cells[i, 1].Value = lastWeekDate.AddDays(i - 1).ToShortDateString();
                sheet.Cells[i + 7, 1].Value = lastWeekDate.AddDays(i + 7 - 1).ToShortDateString();
                sheet.Cells[i + 14, 1].Value = lastWeekDate.AddDays(i + 14 - 1).ToShortDateString();

                sheet.Cells[i, 2].Value = lastWeekDate.AddDays(i - 1).ToShortDateString();
                sheet.Cells[i + 7, 2].Value = lastWeekDate.AddDays(i + 7 - 1).ToShortDateString();
                sheet.Cells[i + 14, 2].Value = lastWeekDate.AddDays(i + 14 - 1).ToShortDateString();

                sheet.Cells[i, 3].Value = lastWeekDate.AddDays(i - 1).ToShortDateString();
                sheet.Cells[i + 7, 3].Value = lastWeekDate.AddDays(i + 7 - 1).ToShortDateString();
                sheet.Cells[i + 14, 3].Value = lastWeekDate.AddDays(i + 14 - 1).ToShortDateString();

                sheet.Cells[i, 5].Value = lastMonth + $"{i + 10}";
                sheet.Cells[i + 7, 5].Value = thisMonth + $"{i + 10}";
                sheet.Cells[i + 14, 5].Value = nextMonth + $"{i + 10}";

                sheet.Cells[i, 6].Value = lastMonth + $"{i + 10}";
                sheet.Cells[i + 7, 6].Value = thisMonth + $"{i + 10}";
                sheet.Cells[i + 14, 6].Value = nextMonth + $"{i + 10}";

                sheet.Cells[i, 7].Value = lastMonth + $"{i + 10}";
                sheet.Cells[i + 7, 7].Value = thisMonth + $"{i + 10}";
                sheet.Cells[i + 14, 7].Value = nextMonth + $"{i + 10}";
            }
        }
    }
}
