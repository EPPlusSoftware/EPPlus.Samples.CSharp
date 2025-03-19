using EPPlusSamples.DrawingsChartsAndThemes;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples._05_Drawings_charts_and_themes._03_Charts_and_themes
{
    public class ShapesAndPicturesInChartSample : ChartSampleBase
    {
        public static async Task Add(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Chart Drawings");

            var fullRange = await LoadFromDatabase(ws);
            var range = fullRange.SkipRows(1);

            //Add a line chart
            var lineChart = ws.Drawings.AddLineChart("line3dChart", eLineChartType.Line3D);
            var lineSerie = lineChart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
            lineSerie.Header = "Order Value";
            lineChart.SetPosition(21, 0, 6, 0);
            lineChart.SetSize(1200, 400);
            lineChart.Title.Text = "Line 3D";

            //Add a shape to the chart.
            var circle = lineChart.Drawings.AddShape("Circle", eShapeStyle.Ellipse);
            //Position and resize the shape.
            circle.SetPosition(144, 258);
            circle.SetSize(120);
            circle.Fill.Color = Color.Orange;

            //Add a picture to the chart.
            var pic = lineChart.Drawings.AddPicture("Logo", FileUtil.GetFileInfo("05-Drawings charts and themes\\01-Shapes and images", "EPPlusLogo.jpg"));
            //Position the picture.
            pic.SetPosition(200, 200);
        }
    }
}
