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
    internal class StylingExample
    {
        public static void Run(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("StyleExample");

            //Below applies to almost every ConditionalFormatting except Databars, Iconsets and ColorScales.

            //They work similarily to ordinary cell styles with a few restrictions.
            //ConditionalFormattings (CFs) have 4 major style categories: Fill, Border, Font and Numberformat.
            //Each roughly equivalent to a formatting tab in excel.

            sheet.Cells["A1:A10"].Formula = "ROW()";

            var cf = sheet.ConditionalFormatting.AddAboveAverage("A1:A10");

            //Fill which defines the inside of cells.
            //Its first important property is style which defines if it follows a pattern or a gradient.
            //Gradient is accessed under fillEffects in excel but epplus provides a shorthand.
            cf.Style.Fill.Style = eDxfFillStyle.PatternFill;

            //The most common type of fill "Solid Fill" is a pattern fill.
            //This property represents the Pattern Style drop down in excel and has enum options for all of them.
            cf.Style.Fill.PatternType = ExcelFillStyle.Solid;

            //This is how to pick "thin horizontal" equivalent in excel
            //Note that the name is as it needs to be written in the xml.
            cf.Style.Fill.PatternType = ExcelFillStyle.LightVertical;

            //Represents Pattern Color in excel .Gradient is the equivalent for gradient styles.
            cf.Style.Fill.BackgroundColor.Color = Color.Gold;

            //.Border refers to the borders around a cell. You can set different options for each of the four or .BorderAround for all borders
            cf.Style.Border.Top.Style = ExcelBorderStyle.MediumDashed;
            cf.Style.Border.Top.Color.Color = Color.RebeccaPurple;

            //This will overwrite the previous changes but also apply to all borders
            cf.Style.Border.BorderAround(ExcelBorderStyle.MediumDashDotDot, Color.Red);

            //.Font has multiple standard properties like the below
            cf.Style.Font.Bold = true;
            cf.Style.Font.Italic = true;
            cf.Style.Font.Strike = true;
            cf.Style.Font.Underline = ExcelUnderLineType.Single;
            cf.Style.Font.Color.Color = Color.ForestGreen;

            //NumberFormat represents the Number tab of the format UI in excel and is set via format string
            cf.Style.NumberFormat.Format = "0.00%";
            //You can also get the id of the numberformat but not set it
            var id = cf.Style.NumberFormat.NumFmtID;

            //Note that this worksheet will look strange as we add a lot of options on just one formatting.
        }
    }
}
