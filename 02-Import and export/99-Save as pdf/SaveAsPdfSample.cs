using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples._02_Import_and_export._99_Save_as_pdf
{
    public static class SaveAsPdfSample
    {
        //This sample demonstrates how to save workbook, worksheets and ranges to pdf.
        public static void Run()
        {
            Console.WriteLine("Running sample 99 - Save as pdf");
            var outputFolder = FileUtil.GetDirectoryInfo("PdfOutput").ToString();
            //Start by using the excel file generated in sample 28
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("Workbooks", "2.5-LoadingData.xlsx")))
            {
                var workbook = p.Workbook;
                var worksheet1 = workbook.Worksheets[0];
                var worksheet2 = workbook.Worksheets[0];
                var worksheet3 = workbook.Worksheets[0];

                ExportWorksheet(outputFolder, worksheet1);

                ExportWorkbook(outputFolder, workbook);

                ExportWorksheets(outputFolder, workbook, worksheet1, worksheet3);

                ExportRange(outputFolder, worksheet3);

                ExportRanges(outputFolder, workbook, worksheet3);
            }
        }

        public static void ExportWorksheet(string outputFolder, ExcelWorksheet worksheet)
        {
            //Use the printer settings on the worksheet for options.
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            //You can set header and footer using HeaderFooter on the worksheet.
            worksheet.HeaderFooter.OddHeader.LeftAlignedText = "Sample pdf";
            //The worksheet is saved as pdf.
            worksheet.SaveAsPdf(outputFolder + "Sample 99.1.pdf");
        }


        public static void ExportWorkbook(string outputFolder, ExcelWorkbook workbook)
        {
            //Currently only the first worksheets printer settings is used for the whole workbook(This will change in a future release).
            workbook.Worksheets[0].PrinterSettings.PaperSize = ePaperSize.A3;
            workbook.Worksheets[0].PrinterSettings.Orientation = eOrientation.Landscape;
            //The workbook is saved as pdf.
            workbook.SaveAsPdf(outputFolder + "Sample 99.2.pdf");
        }

        public static void ExportWorksheets(string outputFolder, ExcelWorkbook workbook, ExcelWorksheet worksheet1, ExcelWorksheet worksheet2)
        {
            //Setup printer settings on the first worksheet.
            worksheet1.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet1.PrinterSettings.ShowGridLines = true;
            worksheet1.PrinterSettings.ShowHeaders = true;
            //Using the SaveAsPdf on the workbook allows multiple worksheets.
            workbook.SaveAsPdf(outputFolder + "Sample 99.3.pdf", worksheet1, worksheet2);
        }

        public static void ExportRange(string outputFolder, ExcelWorksheet worksheet)
        {
            //Setup printer settings on the worksheet.
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet.PrinterSettings.ShowGridLines = true;
            worksheet.PrinterSettings.ShowHeaders = true;
            //Get the range to export.
            var range = worksheet.Cells["A1:D42"];
            //Use the save on the range.
            range.SaveAsPdf(outputFolder + "Sample 99.4.pdf");
        }

        public static void ExportRanges(string outputFolder, ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            //Setup printer settings on the worksheet.
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet.PrinterSettings.ShowGridLines = true;
            worksheet.PrinterSettings.ShowHeaders = true;
            //Get the ranges to export.
            var range1 = worksheet.Cells["F1:H4"];
            var range2 = worksheet.Cells["J1:M45"];
            //Using the SaveAsPdf on the workbook allows multiple ranges.
            workbook.SaveAsPdf(outputFolder + "Sample 99.3.pdf", range1, range2);
        }


        //This sample demonstrates how to save workbook, worksheets and ranges to pdf.
        public static async Task RunAsync()
        {
            Console.WriteLine("Running sample 99 - Save as pdf");
            var outputFolder = FileUtil.GetDirectoryInfo("PdfOutput");
            //Start by using the excel file generated in sample 28
            await using (var p = new ExcelPackage(FileUtil.GetFileInfo("Workbooks", "2.5-LoadingData.xlsx")))
            {
                var workbook = p.Workbook;
                var worksheet1 = workbook.Worksheets[0];
                var worksheet2 = workbook.Worksheets[0];
                var worksheet3 = workbook.Worksheets[0];
                await ExportWorksheetAsync(outputFolder, worksheet1);
                await ExportWorkbookAsync(outputFolder, workbook);
                await ExportWorksheetsAsync(outputFolder, workbook, worksheet1, worksheet3);
                await ExportRangeAsync(outputFolder, worksheet3);
                await ExportRangesAsync(outputFolder, workbook, worksheet3);
            }
        }

        public static async Task ExportWorksheetAsync(string outputFolder, ExcelWorksheet worksheet)
        {
            //Use the printer settings on the worksheet for options.
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            //You can set header and footer using HeaderFooter on the worksheet.
            worksheet.HeaderFooter.OddHeader.LeftAlignedText = "Sample pdf";
            //The worksheet is saved as pdf.
            await worksheet.SaveAsPdfAsync(outputFolder + "Sample 99.1.pdf");
        }

        public static async Task ExportWorkbookAsync(string outputFolder, ExcelWorkbook workbook)
        {
            //Currently only the first worksheets printer settings is used for the whole workbook(This will change in a future release).
            workbook.Worksheets[0].PrinterSettings.PaperSize = ePaperSize.A3;
            workbook.Worksheets[0].PrinterSettings.Orientation = eOrientation.Landscape;
            //The workbook is saved as pdf.
            await workbook.SaveAsPdfAsync(outputFolder + "Sample 99.2.pdf");
        }

        public static async Task ExportWorksheetsAsync(string outputFolder, ExcelWorkbook workbook, ExcelWorksheet worksheet1, ExcelWorksheet worksheet2)
        {
            //Setup printer settings on the first worksheet.
            worksheet1.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet1.PrinterSettings.ShowGridLines = true;
            worksheet1.PrinterSettings.ShowHeaders = true;
            //Using the SaveAsPdfAsync on the workbook allows multiple worksheets.
            await workbook.SaveAsPdfAsync(outputFolder + "Sample 99.3.pdf", worksheet1, worksheet2);
        }

        public static async Task ExportRangeAsync(string outputFolder, ExcelWorksheet worksheet)
        {
            //Setup printer settings on the worksheet.
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet.PrinterSettings.ShowGridLines = true;
            worksheet.PrinterSettings.ShowHeaders = true;
            //Get the range to export.
            var range = worksheet.Cells["A1:D42"];
            //Use the save on the range.
            await range.SaveAsPdfAsync(outputFolder + "Sample 99.4.pdf");
        }

        public static async Task ExportRangesAsync(string outputFolder, ExcelWorkbook workbook, ExcelWorksheet worksheet)
        {
            //Setup printer settings on the worksheet.
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet.PrinterSettings.ShowGridLines = true;
            worksheet.PrinterSettings.ShowHeaders = true;
            //Get the ranges to export.
            var range1 = worksheet.Cells["F1:H4"];
            var range2 = worksheet.Cells["J1:M45"];
            //Using the SaveAsPdfAsync on the workbook allows multiple ranges.
            await workbook.SaveAsPdfAsync(outputFolder + "Sample 99.3.pdf", range1, range2);
        }
    }
}
