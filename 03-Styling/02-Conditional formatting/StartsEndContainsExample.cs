using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.ConditionalFormatting
{
    internal class StartsEndContainsExample
    {
        public static void Run(ExcelPackage package) 
        {
            var sheet = package.Workbook.Worksheets.Add("StartEndContains");

            sheet.Cells["E11"].Value = "SearchMe will find me but the notContains will not. Nope NotMe never.";
            sheet.Cells["E12"].Value = "SearchMe will find me but not EndText because I'm a EndTextFaker";
            sheet.Cells["E13"].Value = "NotMe won't be found by much";
            sheet.Cells["E14"].Value = "This will be found by notContains and is also EndText";
            sheet.Cells["E15"].Value = "This will be found by notContains and is also EndText";
            sheet.Cells["E16"].Value = "SearchMe To be found by all and let the end be EndText";

            // -------------------------------------------------------------------
            // Create a BeginsWith rule
            // -------------------------------------------------------------------
            ExcelAddress cellIsAddress = new ExcelAddress("E11:E20");
            var beginsWith = sheet.ConditionalFormatting.AddBeginsWith(
              cellIsAddress);

            beginsWith.Text = "SearchMe";

            beginsWith.Style.Font.Bold = true;

            // -------------------------------------------------------------------
            // Create an EndsWith rule
            // -------------------------------------------------------------------
            var EndText = sheet.ConditionalFormatting.AddEndsWith(
              cellIsAddress);

            EndText.Text = "EndText";

            EndText.Style.Font.Color.Color = Color.DarkRed;

            // -------------------------------------------------------------------
            // Create a ContainsText rule
            // -------------------------------------------------------------------
            var ContainsText = sheet.ConditionalFormatting.AddContainsText(
              cellIsAddress);

            ContainsText.Text = "Me";
            ContainsText.Style.Font.Italic = true;

            // -------------------------------------------------------------------
            // Create a NotContainsText rule
            // -------------------------------------------------------------------
            var notContains = sheet.ConditionalFormatting.AddNotContainsText(
              cellIsAddress);

            notContains.Text = "NotMe";
            notContains.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.MediumDashed, Color.Red);
        }
    }
}
