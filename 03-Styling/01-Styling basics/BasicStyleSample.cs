using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.Styling
{
    public static class BasicStyleSample
    {
        public static void Run()
        {
            FileInfo newFile = FileUtil.GetCleanFileInfo("3.1-Basic Styling.xlsx");
            using (var package = new ExcelPackage(newFile))
            {
                //Formatting cells
                Style_NumberFormats(package);
                Style_FontAndFill(package);
                Style_Borders(package);
                Style_Alignments(package);
                package.Save();
            }
        }

        private static void Style_Alignments(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Alignments");
            ws.Columns[1].Width = 15;
            ws.Cells["A1:A6"].Value = "Test Text Styling";
            ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            ws.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            ws.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["A4:C4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Justify;
            ws.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left; //Left indented alignment
            ws.Cells["A6"].Style.Indent = 2;

            ws.Cells["A10:E10"].Value = "Test of text alignment";
            ws.Rows[10].Height = 60;
            ws.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            ws.Cells["B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ws.Cells["C10"].Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
            ws.Cells["D10"].Style.VerticalAlignment = ExcelVerticalAlignment.Justify;
            ws.Cells["E10"].Style.SetTextVertical();
        }

        private static void Style_Borders(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Border");
            ws.Cells["B2"].Style.Border.BorderAround(ExcelBorderStyle.Dashed);
            ws.Cells["B2"].Value = "Dashed Border";            
            
            ws.Cells["B4:C5"].Style.Border.BorderAround(ExcelBorderStyle.Dotted, Color.Black);            
            ws.Cells["B4:C5"].Merge = true;
            ws.Cells["B4:C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ws.Cells["B4:C5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["B4"].Value = "Border Around Merged";

            var rB7 = ws.Cells["B7"];
            ws.Cells["B7"].Value = "Mixed borders";
            rB7.Style.Border.Top.Style = ExcelBorderStyle.MediumDashed;
            rB7.Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent1);
            
            rB7.Style.Border.Bottom.Style = ExcelBorderStyle.DashDotDot;
            rB7.Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent2);

            rB7.Style.Border.Left.Style = ExcelBorderStyle.Double;
            rB7.Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent3);

            rB7.Style.Border.Right.Style = ExcelBorderStyle.Thick;
            rB7.Style.Border.Top.Color.SetColor(eThemeSchemeColor.Accent4);

            ws.Cells.AutoFitColumns();
        }

        private static void Style_FontAndFill(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Fonts & Fills");

            //Font Styles
            ws.Cells["A1:A11"].Value = "Font";
            ws.Cells["A1"].Style.Font.Bold = true;
            ws.Cells["A2"].Style.Font.Italic = true;
            ws.Cells["A3"].Style.Font.UnderLine = true;
            ws.Cells["A4"].Style.Font.UnderLineType = ExcelUnderLineType.Double;
            ws.Cells["A5"].Style.Font.Strike = true;
            ws.Cells["A6"].Style.Font.Color.SetColor(Color.DarkRed);
            ws.Cells["A7"].Style.Font.Color.SetColor(eThemeSchemeColor.Text2);
            ws.Cells["A8"].Style.Font.Color.SetColor(ExcelIndexedColor.Indexed3);
            ws.Cells["A9"].Style.Font.Size=18;
            ws.Cells["A10"].Style.Font.Name = "Arial";
            ws.Cells["A11"].Style.Font.VerticalAlign = ExcelVerticalAlignmentFont.Subscript;

            //Cell Fills
            ws.Cells["B1:B11"].Value = "Fills";
            ws.Cells["B1"].Style.Fill.SetBackground(Color.LightGray);
            ws.Cells["B2"].Style.Fill.SetBackground(Color.Gray, ExcelFillStyle.DarkGrid);

            var rB3 = ws.Cells["B3"];
            rB3.Style.Fill.PatternType = ExcelFillStyle.DarkDown;
            rB3.Style.Fill.PatternColor.SetColor(Color.DarkRed);
            rB3.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);

            var rB4 = ws.Cells["B4"];
            rB4.Style.Fill.Gradient.Color1.SetColor(Color.Green);
            rB4.Style.Fill.Gradient.Color2.SetColor(Color.Yellow);
            rB4.Style.Fill.Gradient.Degree = 90;
            rB4.Style.Fill.Gradient.Top = 0.8;
            rB4.Style.Fill.Gradient.Bottom = 0.8;
            rB4.Style.Fill.Gradient.Left = 0.8;
            rB4.Style.Fill.Gradient.Right = 0.8;

            var rB5 = ws.Cells["B5"];
            rB4.Style.Fill.Gradient.Color2.SetColor(Color.Yellow);
            rB4.Style.Fill.Gradient.Degree = 90;
            rB4.Style.Fill.Gradient.Top = 0.8;
            rB4.Style.Fill.Gradient.Bottom = 0.8;
            rB4.Style.Fill.Gradient.Left = 0.8;
            rB4.Style.Fill.Gradient.Right = 0.8;
        }

        private static void Style_NumberFormats(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Numberformats");
            ws.Cells["A1"].Value = "Numbers formats";
            ws.Cells["A10"].Value = "Other formats";

            ws.Cells["A1:E1"].Merge = true;
            ws.Cells["A10:E10"].Merge = true;
            ws.Cells["A1,A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
            ws.Cells["A1,A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            ws.Cells["A2:E2"].Value = 5;
            ws.Cells["A3:E3"].Value = 1.23E+15;
            ws.Cells["A4:E4"].Value = 3.21547218012E-3;
            ws.Cells["A5:E5"].Value = -123456.321;
            ws.Cells["A6:E6"].Value = 0.1221;
            ws.Cells["A7:E7"].Value = 0;
            ws.Cells["A8:E8"].Value = "11";

            ws.Cells["A2:A9"].Style.Numberformat.Format = "#,##0.00";                  // Numberformat with 1000-separator. This format correspons to the buildin format 4. See https://github.com/EPPlusSoftware/EPPlus/wiki/Formatting-and-styling#number-formats
            ws.Cells["B2:B9"].Style.Numberformat.Format = "[Green]#,##0.00;[Red] (#,##0.00);[Blue]#,##0.00;[Cyan]@"; // Format with some different font colors and parantheses for negative values. First part is positiv, then negative, the zero and last part is for Text. You can use [black] [white] [red][green] [blue] [yellow] [magenta] and [cyan]
            ws.Cells["C2:C9"].Style.Numberformat.Format = "[$$-1009]#,##0.00";         // Localized currency to Canadian format. The LCID in hex for the CultureInfo is specified inside the brackets. 
            ws.Cells["D2:D9"].Style.Numberformat.Format = "0.000%";                    // Number format with percent.
            ws.Cells["E2:E9"].Style.Numberformat.Format = "@";                         // Text.

            ws.Cells["A11:E11"].Value = new DateTime(DateTime.Today.Year, 1, 1);
            ws.Cells["A12:E12"].Value = new DateTime(DateTime.Today.Year, 3, 31, 12, 30, 35);
            ws.Cells["A13:E13"].Value = DateTime.Parse("13:45:00", CultureInfo.InvariantCulture);
            ws.Cells["A14:E14"].Value = new TimeSpan(6, 30, 0);

            ws.Cells["A11:A19"].Style.Numberformat.Format = "yyyy-MM-dd";                               //Sets the number format to short date format with regional formatting
            ws.Cells["B11:B19"].Style.Numberformat.Format = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";    //Sets the number format to long date format with regional formatting. F800 specifies that the long format should be used. 
            ws.Cells["C11:C19"].Style.Numberformat.Format = "[$-F400]h:mm:ss\\ am/pm";                  //Sets the number format to long time format with regional formatting. F400 specifies that the long format should be used. 
            ws.Cells["D11:D19"].Style.Numberformat.Format = "[$]hh:mm:ss;@";                            //Short time format. Last part is for cells containing text.
            ws.Cells["E11:E19"].Style.Numberformat.Format = "[$-409]h:mm\\ am/pm;@";                    //Format time using culture LCID 0x409 (1033).

            ws.Cells.AutoFitColumns();
        }
    }
}
