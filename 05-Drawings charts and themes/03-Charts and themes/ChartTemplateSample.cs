using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System.Threading.Tasks;

namespace EPPlusSamples.DrawingsChartsAndThemes
{
    public class ChartTemplateSample : ChartSampleBase
    {
        public static async Task AddAreaChart(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Area chart from template");
            var range = await LoadFromDatabase(ws);

            //Adds an Area chart from a template file. The crtx file has it's own theme, so it does not change with the theme.
            //The As property provides an easy type cast for drawing objects
            var areaChart = ws.Drawings.AddChartFromTemplate(FileUtil.GetFileInfo("05-Drawings charts and themes\\03-Charts and themes", "AreaChartStyle3.crtx"), "areaChart")
                .As.Chart.AreaChart;
            var areaSerie = areaChart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            areaSerie.Header = "Order Value";
            areaChart.SetPosition(1, 0, 6, 0);
            areaChart.SetSize(1200, 400);
            areaChart.Title.Text = "Area Chart";

            range.AutoFitColumns(0);
        }
    }
}
