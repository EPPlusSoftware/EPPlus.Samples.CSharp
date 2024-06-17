using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations.Model;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing.Chart;
using System.Data;
using System.Drawing;
using OfficeOpenXml.Table;

namespace EPPlusSamples._05_Drawings_charts_and_themes._03_Charts_and_themes
{
    public class BarColumnChartsWithManualLayout
    {
        public static void Add(ExcelPackage package)
        {
            var cSheet = package.Workbook.Worksheets.Add("ColumnChartSheet");

            var range = cSheet.Cells["A1:C3"];
            var table = cSheet.Tables.Add(range, "DataTable");
            table.ShowHeader = false;
            table.SyncColumnNames(ApplyDataFrom.ColumnNamesToCells, true);

            range.Formula = "ROW() + COLUMN()";

            cSheet.Calculate();

            var sChart = cSheet.Drawings.AddBarChart("simpleChart", eBarChartType.ColumnStacked);

            sChart.Series.Add(cSheet.Cells["A1:A3"]);
            sChart.Series.Add(cSheet.Cells["B1:B3"]);
            sChart.Series.Add(cSheet.Cells["C1:C3"]);

            var highestSeriesOfColumns = sChart.Series[2];

            var dataLabelRulesForEntireRow = highestSeriesOfColumns.DataLabel;

            var topColumnInStack = dataLabelRulesForEntireRow.DataLabels.Add(0);
            topColumnInStack.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            topColumnInStack.Fill.SolidFill.Color.SetRgbColor(Color.MediumPurple);

            var middleColumnInSecondStack = sChart.Series[1].DataLabel.DataLabels.Add(1);

            var bottomColumnInFirstStack = sChart.Series[0].DataLabel.DataLabels.Add(0);
            bottomColumnInFirstStack.Fill.Style = OfficeOpenXml.Drawing.eFillStyle.SolidFill;
            bottomColumnInFirstStack.Fill.SolidFill.Color.SetRgbColor(Color.CornflowerBlue);

            var lastTop = sChart.Series[2].DataLabel.DataLabels.Add(2);

            SetShowValues(topColumnInStack);
            SetShowValues(middleColumnInSecondStack);
            SetShowValues(bottomColumnInFirstStack);
            SetShowValues(lastTop);

            var manualLayoutTop = topColumnInStack.Layout.ManualLayout;
            var middleColumnLayout = middleColumnInSecondStack.Layout.ManualLayout;
            var bottomColumnLayout = bottomColumnInFirstStack.Layout.ManualLayout;
            var lastTopLayout = lastTop.Layout.ManualLayout;

            //Set x and y position (units are in percent of chart width/height)
            //Left means pushing 'from' the left. Same with top. It's the position of the left side of the element box.
            manualLayoutTop.Left = -5;
            manualLayoutTop.Top = -25;

            //Set textbox width
            manualLayoutTop.Width = 5;
            manualLayoutTop.Height = 10;

            //Set x only
            middleColumnLayout.Left = 10;

            ////By default left and top are offsets to the starting Position of the element. AKA: dataLabel.Position
            ////To make positioning easier you can also define starting position as an offset from the Left or Top edge of the chart:
            bottomColumnLayout.TopMode = eLayoutMode.Edge;
            bottomColumnLayout.LeftMode = eLayoutMode.Edge;

            //This will put the label in the top left corner of the chart itself.
            bottomColumnLayout.Left = 0;
            bottomColumnLayout.Top = 0;

            //Note that when in edge mode, negative inputs are nonsensical as 0 is the starting point and negative values would be outside the chart.
            //Forcing a label outside the chart resets it to its Position attribute in excel.
            lastTopLayout.TopMode = eLayoutMode.Edge;
            lastTopLayout.Top = -5;
        }

        private static void SetShowValues(ExcelChartDataLabelItem item)
        {
            item.ShowLegendKey = false;
            item.ShowValue = true;
            item.ShowCategory = false;
            item.ShowSeriesName = false;
            item.ShowPercent = false;
            item.ShowBubbleSize = false;
            item.ShowLeaderLines = true;
        }
    }
}
