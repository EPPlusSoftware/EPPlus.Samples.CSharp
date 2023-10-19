using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class DatabarsExample
    {
        public static void Run(ExcelPackage pck)
        {
            var ws = pck.Workbook.Worksheets.Add("Databars");

            ws.Cells["A2:H21"].Formula = "Row() - 11";

            //Adding gradient databar
            ws.Cells["A2:A21"].ConditionalFormatting.AddDatabar(Color.OrangeRed);
            ws.Cells["A1"].Value = "Gradient";

            //Solid Color Databar
            var databar = ws.Cells["B2:B21"].ConditionalFormatting.AddDatabar(Color.BlueViolet);
            databar.Gradient = false;

            ws.Cells["B1"].Value = "Solid";

            //Below only accesible epplus7 and beyond

            //Themecolor bar note that input color does not matter if fill is changed
            var databarTheme = ws.Cells["C2:C21"].ConditionalFormatting.AddDatabar(Color.BlueViolet);

            ws.Cells["C1"].Value = "ThemeColor";

            databarTheme.FillColor.SetColor(eThemeSchemeColor.Accent2);
            //You can also set the border color
            databarTheme.BorderColor.SetColor(Color.Green);

            //Auto color
            var databarAuto = ws.Cells["D2:D21"].ConditionalFormatting.AddDatabar(Color.Red);
            ws.Cells["D1"].Value = "AutoColor";

            //Note: Auto color is white
            databarAuto.FillColor.SetAuto();
            //Making the white visible by filling a background color
            ws.Cells["D10:D21"].Style.Fill.PatternType = ExcelFillStyle.Solid;
            ws.Cells["D10:D21"].Style.Fill.BackgroundColor.SetColor(Color.Cornsilk);


            //Indexed color (excel legacy)
            var databarIndexed = ws.Cells["E2:E21"].ConditionalFormatting.AddDatabar(Color.Red);
            ws.Cells["E1"].Value = "IndexAndNegativeColors";

            databarIndexed.FillColor.SetColor(ExcelIndexedColor.Indexed12);

            //similarily you can also apply all these settings to negative bar colors and borders
            databarIndexed.NegativeFillColor.SetColor(eThemeSchemeColor.Accent4);
            databarIndexed.NegativeBorderColor.SetColor(ExcelIndexedColor.Indexed45);
            //And the axis between negative and positive numbers
            databarIndexed.AxisColor.SetColor(Color.Purple);

            //Alternatively positive and negative colors can just be the same
            var boolsEx = ws.Cells["F2:F21"].ConditionalFormatting.AddDatabar(Color.Green);
            ws.Cells["F1"].Value = "SameAsPositive";

            boolsEx.BorderColor.SetColor(Color.Black);

            boolsEx.NegativeBarBorderColorSameAsPositive = true;
            boolsEx.NegativeBarColorSameAsPositive = true;

            //--------------------------------------------------------
            //Databars also contain other settings such as these

            var dataBarWithSettings = ws.Cells["G2:G21"].ConditionalFormatting.AddDatabar(Color.Blue);
            ws.Cells["G1"].Value = "MultipleSettings";

            dataBarWithSettings.AxisColor.SetColor(Color.Purple);
                
            dataBarWithSettings.AxisPosition = OfficeOpenXml.ConditionalFormatting.eExcelDatabarAxisPosition.Automatic;
            //Direction of the databar (Default is left to right)
            dataBarWithSettings.Direction = OfficeOpenXml.ConditionalFormatting.eDatabarDirection.RightToLeft;

            //Define when the databars length reaches its maximum and minimum value
            dataBarWithSettings.HighValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Num;
            dataBarWithSettings.LowValue.Type = OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType.Num;

            dataBarWithSettings.HighValue.Value = 5;
            dataBarWithSettings.LowValue.Value = -5;

            var dbSameDirection = ws.ConditionalFormatting.AddDatabar("H2:H21", Color.Yellow);
            ws.Cells["H1"].Value = "SameDirection";

            //Show negative and positive bars in same direction
            dbSameDirection.AxisPosition = OfficeOpenXml.ConditionalFormatting.eExcelDatabarAxisPosition.None;
            ws.Cells.AutoFitColumns();
        }
    }
}
