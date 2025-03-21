/*************************************************************************************************
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
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.OleObject;
using System.IO;

namespace EPPlusSamples._05_Drawings_charts_and_themes._06_OLE_Objects
{
    /// <summary>
    /// This sample shows how to embed or link files as OLE Objects using EPPLUS.
    /// </summary>
    public static class OLEObjectsSample
    {
        public static void Run()
        {
            var myPDF = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "MyPDF.pdf");
            var myWord = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "MyWord.docx");
            var myTxt = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "MyTextDocument.txt");
            var myIcon = FileUtil.GetFileInfo("05-Drawings charts and themes\\06-OLE Objects", "SampleIcon.bmp");
            FileInfo newWorkbook = FileUtil.GetCleanFileInfo(@"5.6-OLE Objects.xlsx");


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
            var LinkedPDF = ws2.Drawings.AddOleObject("MyPDF", myPDF, o => o.LinkToFile = true);

            //Save the workbook
            p2.SaveAs(newWorkbook);

            /* Link a file with ExcelOleObjectParameters.    */
            // Create a workbook and a worksheet.
            using var p3 = new ExcelPackage(newWorkbook);
            var ws3 = p3.Workbook.Worksheets.Add("Sheet 3");
            //Link the file using AddOleObject method on the drawing with additional parameters.
            var LinkedPDF2 = ws3.Drawings.AddOleObject("MyPDF", myPDF, o =>
            {
                o.DisplayAsIcon = true;
                o.LinkToFile = true;
            });

            //Save the workbook
            p3.SaveAs(newWorkbook);


            /*    Add custom icon.    */
            //Create a workbook and a worksheet.
            using var p4 = new ExcelPackage(newWorkbook);
            var ws4 = p4.Workbook.Worksheets.Add("Sheet 4");
            //Link the file using AddOleObject method on the drawing.
            var txt = ws4.Drawings.AddOleObject("MyText", myTxt, o =>
            {
                o.DisplayAsIcon = true;
                o.LinkToFile = true;
                o.Icon = new ExcelImage(myIcon);
            });
            //Save the workbook
            p4.SaveAs(newWorkbook);


            /*    Copy OLE Object    */
            //Create a workbook, get the worksheet and create a new worksheet.
            using var p5 = new ExcelPackage(newWorkbook);
            var ws1 = p5.Workbook.Worksheets[0];
            var ws5 = p5.Workbook.Worksheets.Add("Sheet 5");
            //Get OLE Object.
            var newPDF = ws1.Drawings[0] as ExcelOleObject;
            //Copy OLE Object to a new worksheet.
            var copy = newPDF.Copy(ws5, 1, 4);
            //Save the workbook
            p5.SaveAs(newWorkbook);


            /*    Delete OLE Object    */
            //Create a workbook and get worksheet.
            using var p6 = new ExcelPackage(newWorkbook);
            var ws6 = p6.Workbook.Worksheets.Add("Delete OLE object");
            //Get the OLE Object from worksheet 1.
            var myPdfDoc = ws1.Drawings[0] as ExcelOleObject;

            //Copy the OLE Object to a new worksheet.
            var copyToDelete = myPdfDoc.Copy(ws6, 1, 4);
            //Now remove the OLE Object from the worksheet.
            ws6.Drawings.Remove(copyToDelete);

            //Save the workbook
            p6.SaveAs(newWorkbook);


            /*    Create OLE Object using a stream */
            //Create a workbook and create a new worksheet.
            using var p7 = new ExcelPackage(newWorkbook);
            var ws7 = p7.Workbook.Worksheets.Add("Sheet 7");
            //Create the stream
            using (FileStream fileStream = new FileStream(myPDF.FullName, FileMode.Open, FileAccess.Read))
            {
                //Add OLE Object using stream and filename.
                var oleFromStream = ws7.Drawings.AddOleObject("MyPdfFromStream", fileStream, "MyPdfFromStream.pdf");
            }
            p7.SaveAs(newWorkbook);
        }
    }
}
