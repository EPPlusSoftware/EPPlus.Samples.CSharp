/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/29/2024         EPPlus Software AB           Initial release EPPlus 7
 *************************************************************************************************/
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace EPPlusSamples._02_Import_and_export._03_Import_export_text_files
{
    /// <summary>
    /// This sample shows how to load/save Fixed Width files using the LoadFromText and SaveToText methods.
    /// </summary>
    public static class Sample020302_ImportAndExportFixedWidthFiles
    {
        public static void RunSample()
        {
            Console.WriteLine("Running sample 2.3.2");
            var fixedWidthFile = FileUtil.GetFileInfo("02-Import and Export\\03-Import export text files", "Sample2.3.2-1.txt");
            FileInfo newFile = FileUtil.GetCleanFileInfo(@"2.3.2-LoadDataFromFixedWidthFiles.xlsx");

            //Import fixed width text file using column length.
            {
                Console.WriteLine("Importing the file using column lengths...");
                //Create a workbook and a worksheet.
                using var package = new ExcelPackage();
                var sheet = package.Workbook.Worksheets.Add("FixedWidthLengths");

                //Create the import settings object.
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();

                //Set the length of each column.
                format.SetColumnLengths(16, 10, 16, 8, 1);
                //Skip the header row.
                format.SkipLinesBeginning = 1;

                //Load the fixed width text file into.
                var range = sheet.Cells["A1"].LoadFromText(fixedWidthFile, format);

                //Save the excel file.
                package.SaveAs(newFile);
            }





            //Import fixed width text file using column starting position.
            {
                Console.WriteLine("Importing the file using column positions...");
                //Create a workbook and a worksheet.
                using var package = new ExcelPackage();
                var sheet = package.Workbook.Worksheets.Add("FixedWidthPositions");

                //Create the import settings object.
                ExcelTextFormatFixedWidth format = new ExcelTextFormatFixedWidth();

                //Set the length of a row and the starting positions of each column.
                format.SetColumnPositions(51, 0, 16, 26, 42, 50);
                //Skip the header row.
                format.SkipLinesBeginning = 1;

                //Load the fixed width text file into.
                var range = sheet.Cells["A1"].LoadFromText(fixedWidthFile, format);

                //Save the excel file.
                package.SaveAs(newFile);
            }




            //Export fixed width file using column length.
            {
                Console.WriteLine("Exporting the file using column lengths...");
                //Load workbook and worksheet.
                using var package = new ExcelPackage(/*Path*/);
                var sheet = package.Workbook.Worksheets["FixedWidthLengths"];

                //Create the export settings object.
                ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();

                //Set the length of the row and the staring positions of each column
                format.SetColumnLengths(16, 10, 16, 8, 1);
                //Write header
                format.Header = "Name            Date Amount          Percent Category";

                //Export the range to fixed width text file.
                sheet.Cells["A1:E6"].SaveToText(FileUtil.GetCleanFileInfo("2.3.2-ExportedFromEPPlus.txt"), format);
            }




            //Export fixed width file using column starting position.
            {
                Console.WriteLine("Exporting the file using column positions...");
                //Load workbook and worksheet.
                using var package = new ExcelPackage(/*Path*/);
                var sheet = package.Workbook.Worksheets["FixedWidthPositions"];

                //Create the export settings object.
                ExcelOutputTextFormatFixedWidth format = new ExcelOutputTextFormatFixedWidth();

                //Set the length of the row and the staring positions of each column
                format.SetColumnPositions(51, 0, 16, 26, 42, 50);
                //Write header
                format.Header = "Name            Date Amount          Percent Category";

                //Export the range to fixed width text file.
                sheet.Cells["A1:E6"].SaveToText(FileUtil.GetCleanFileInfo("2.3.2-ExportedFromEPPlus.txt"), format);
            }
        }
    }
}
