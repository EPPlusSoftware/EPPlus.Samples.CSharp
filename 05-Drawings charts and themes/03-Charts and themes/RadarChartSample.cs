using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
namespace EPPlusSamples.DrawingsChartsAndThemes
{
    public class RadarChartSample : ChartSampleBase
    {
        public static void Add(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("RadarChart");

            var dt = GetCarDataTable();
            var fullRange = ws.Cells["A1"].LoadFromDataTable(dt, true);
            var range = fullRange.SkipRows(1);
            range.AutoFitColumns();

            var chart = ws.Drawings.AddRadarChart("RadarChart1", eRadarChartType.RadarFilled);
            for(var col = 1; col < fullRange.Columns; col++)
            {
                var serie = chart.Series.Add(range.TakeSingleColumn(col), range.TakeSingleColumn(0));
                serie.HeaderAddress = fullRange.TakeSingleCell(0, col);
            }

            chart.Legend.Position = eLegendPosition.Top;
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.RadarChartStyle4);

            //If you want to apply custom styling do that after setting the chart style so its not overwritten.
            chart.Legend.Effect.SetPresetShadow(ePresetExcelShadowType.OuterTopLeft);

            chart.SetPosition(0, 0, 6, 0);
            chart.To.Column = 17;
            chart.To.Row = 30;
        }

    }
}
