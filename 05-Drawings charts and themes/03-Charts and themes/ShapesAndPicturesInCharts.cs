/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/06/2025         EPPlus Software AB           EPPlus 8.1
 *************************************************************************************************/
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Threading.Tasks;

namespace EPPlusSamples.DrawingsChartsAndThemes
{
    public class ShapesAndPicturesInCharts : ChartSampleBase
    {
        public static async Task Add(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Shape & Picture in Chart");

            var fullRange = await LoadFromDatabase(ws);
            var range = fullRange.SkipRows(1); // remove the headers row

            //Add a line chart
            var chart = ws.Drawings.AddBarChart("BarChartWithPictureAndShape",eBarChartType.PyramidCol);
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);

            var series = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
            series.Header = "Order Value";
            
            var picture = chart.Drawings.AddPicture("logo", FileUtil.GetFileInfo("05-Drawings charts and themes\\03-Charts and themes", "EPPlusLogo.jpg"));
            picture.SetPosition(5, 5);
            picture.SetSize(40);

            var heart = chart.Drawings.AddShape("Heart", eShapeStyle.Heart);            
            
            heart.SetPosition(5, 1140);
            heart.SetSize(40, 40);
            heart.ChangeCellAnchor(eEditAs.Absolute);

            chart.Title.Text = "Pyramid Chart with Shapes";
        }
    }
}