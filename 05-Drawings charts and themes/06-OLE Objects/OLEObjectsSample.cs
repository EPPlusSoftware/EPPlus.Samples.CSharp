﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2025         EPPlus Software AB           Initial release EPPlus 8
 *************************************************************************************************/
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace EPPlusSamples._05_Drawings_charts_and_themes._06_OLE_Objects
{
    /// <summary>
    /// This sample shows how to embed or link files as OLE Objects using EPPLUS.
    /// </summary>
    public static class OLEObjectsSample
    {
        public static void RunSample()
        {
            var myPDF = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "MyPDF.pdf");
            var myWord = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "MyWord.docx");
            var myTxt = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "MyTextDocument.txt");
            var myIcon = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "SampleIcon.bmp");
            FileInfo newWorkbook = FileUtil.GetCleanFileInfo(@"06-OLE Objects Sample Workbook.xlsx");


            /*    Embedding a file.    */
            //Create a workbook and a worksheet.
            using var p = new ExcelPackage();
            var ws = p.Workbook.Worksheets.Add("Sheet 1");
            //Embed the file using AddOleObject method on the drawing.
            var EmbeddedWord  = ws.Drawings.AddOleObject("MyWord", myWord);
            //Save the workbook
            p.SaveAs(newWorkbook);


            /*    Link a file.    */
            //Create a workbook and a worksheet.
            using var p2 = new ExcelPackage(newWorkbook);
            var ws2 = p2.Workbook.Worksheets.Add("Sheet 2");
            //Link the file using AddOleObject method on the drawing.
            var LinkedPDF = ws2.Drawings.AddOleObject("MyPDF", myPDF, true);
            //Save the workbook
            p2.SaveAs(newWorkbook);


            /*    Link a file with ExcelOleObjectParameters.    */
            //Create a workbook and a worksheet.
            using var p3 = new ExcelPackage(newWorkbook);
            var ws3 = p3.Workbook.Worksheets.Add("Sheet 3");
            //Create parameters object.
            var parameters = new ExcelOleObjectParameters()
            {
                LinkToFile = true,
                DisplayAsIcon = true,
                Extension = ".pdf",
            };
            //Link the file using AddOleObject method on the drawing.
            var LinkedPDF2 = ws3.Drawings.AddOleObject("MyPDF", myPDF, parameters);
            //Save the workbook
            p3.SaveAs(newWorkbook);


            /*    Add custom icon.    */
            //Create a workbook and a worksheet.
            using var p4 = new ExcelPackage(newWorkbook);
            var ws4 = p4.Workbook.Worksheets.Add("Sheet 4");
            //Link the file using AddOleObject method on the drawing.
            var txt = ws4.Drawings.AddOleObject("MyText", myTxt, false, true, myIcon);
            //Save the workbook
            p4.SaveAs(newWorkbook);


            /*    Copy OLE Object    */
            //Create a workbook, get worksheet and create new worksheet.
            using var p5 = new ExcelPackage(newWorkbook);
            var ws1 = p5.Workbook.Worksheets[0];
            var ws5 = p5.Workbook.Worksheets.Add("Sheet 5");
            //Get OLE Object.
            var newPDF = ws1.Drawings[0] as ExcelOleObject;
            //Copy OLE Object to a new worksheet.
            var copy = ws1.Drawings.Copy(ws5, 1, 4);
            //Save the workbook
            p5.SaveAs(newWorkbook);


            /*    Delete OLE Object    */
            //Create a workbook and get worksheet.
            using var p6 = new ExcelPackage(newWorkbook);
            ws1 = p6.Workbook.Worksheets[0];
            //Get OLE Object.
            var myPdfDoc = ws1.Drawings[0] as ExcelOleObject;
            //Copy OLE Object to a new worksheet.
            ws1.Drawings.Remove(myPdfDoc);
            //Save the workbook
            p6.SaveAs(newWorkbook);
        }
    }
}
