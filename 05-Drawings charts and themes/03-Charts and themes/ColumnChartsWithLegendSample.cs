using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Drawing;
using System.Threading.Tasks;

namespace EPPlusSamples.DrawingsChartsAndThemes
{
    public class ColumnChartWithLegendSample : ChartSampleBase
    {
        public static async Task Add(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("ColumnCharts");

            var fullRange = await LoadFromDatabase(ws);
            var range = fullRange.SkipRows(1);

            //Add a line chart
            var chart = ws.Drawings.AddBarChart("ColumnChartWithLegend", eBarChartType.ColumnStacked);
            var serie1 = chart.Series.Add(range.TakeSingleColumn(1), range.TakeSingleColumn(0));
            serie1.Header = "Order Value";
            var serie2 = chart.Series.Add(range.TakeSingleColumn(2), range.TakeSingleColumn(0));
            serie2.Header = "Tax";
            var serie3 = chart.Series.Add(range.TakeSingleColumn(3), range.TakeSingleColumn(0));
            serie3.Header = "Freight";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Column chart";

            //Set style 10
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ColumnChartStyle10);

            chart.Legend.Entries[0].Font.Fill.Color = Color.Red;
            chart.Legend.Entries[1].Font.Fill.Color = Color.Green;
            chart.Legend.Entries[2].Deleted = true;

            range.AutoFitColumns(0);
        }
    }
}
