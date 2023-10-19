using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Linq;

namespace EPPlusSamples.DrawingsChartsAndThemes
{
    public class ScatterChartSample : ChartSampleBase
    {
        public static void Add(ExcelPackage package)
        {
            //Adda a scatter chart on the data with one serie per row. 
            var ws = package.Workbook.Worksheets.Add("Scatter Chart");

            var fullRange = CreateIceCreamData(ws);
            var range = fullRange.SkipRows(1);

            var chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter);
            chart.SetPosition(1, 0, 3, 0);
            chart.To.Column = 18;
            chart.To.Row = 20;
            chart.XAxis.Format = "yyyy-mm";
            chart.XAxis.Title.Text = "Period";
            chart.XAxis.MajorGridlines.Width = 1;
            chart.YAxis.Format = "$#,##0";
            chart.YAxis.Title.Text = "Sales";

            chart.Legend.Position = eLegendPosition.Bottom;

            var serie = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
            serie.HeaderAddress = ws.Cells["A1"];
            var tr = serie.TrendLines.Add(eTrendLine.MovingAverage);
            tr.Name = "Icecream Sales-Monthly Average";
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ScatterChartStyle12);
        }
    }
}
