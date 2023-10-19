﻿using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusSamples.DrawingsChartsAndThemes
{
    public class ChartWorksheetSample : ChartSampleBase
    {
        public static void Add(ExcelPackage package)
        {
            ExcelWorksheet wsData = LoadBubbleChartData(package);

            //Add a bubble chart worksheet on the data with one serie per row. 
            var wsChart = package.Workbook.Worksheets.AddChart("Bubble Chart", eChartType.Bubble);
            var chart = wsChart.Chart.As.Chart.BubbleChart;
            for (int row = 2; row <= 7; row++)
            {
                var serie = chart.Series.Add(wsData.Cells[row, 2], wsData.Cells[row, 3], wsData.Cells[row, 4]);
                serie.HeaderAddress = wsData.Cells[row, 1];
            }

            chart.DataLabel.Position = eLabelPosition.Center;
            chart.DataLabel.ShowSeriesName = true;
            chart.DataLabel.ShowBubbleSize = true;
            chart.Title.Text = "Sales per Region";
            chart.XAxis.Title.Text = "Total Sales";
            chart.XAxis.Title.Font.Size = 12;
            chart.XAxis.MajorGridlines.Width = 1;
            chart.YAxis.Title.Text = "Sold Units";
            chart.YAxis.Title.Font.Size = 12;
            chart.Legend.Position = eLegendPosition.Bottom;

            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.BubbleChartStyle10);
        }
    }
}
