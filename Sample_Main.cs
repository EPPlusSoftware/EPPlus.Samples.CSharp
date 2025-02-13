/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using EPPlusSamples.ConditionalFormatting;
using EPPlusSamples.DrawingsChartsAndThemes;
using EPPlusSamples.FormulaCalculation;
using System;
using System.IO;
using System.Threading.Tasks;

namespace EPPlusSamples
{
	class Sample_Main
	{
		static async Task Main(string[] args)
		{
			try
			{
                //EPPlus 5 and later uses a dual license model. This requires you to specifiy the License you are using to be able to use the library. 
                //This sample sets the LicenseContext in the appsettings.json file. An alternative is the commented row below.
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //See https://epplussoftware.com/Developers/LicenseException for more info.

                //Set the output directory to the SampleApp folder where the app is running from. 
                FileUtil.OutputDir = new DirectoryInfo($"{AppDomain.CurrentDomain.BaseDirectory}SampleApp");
                await WorkbookWorksheetAndRangesSamples.RunAsync();
                await ImportAndExportSamples.RunAsync();
                StylingBasics.Run();
                ConditionalFormattingSamples.Run();

                await FiltersAndValidation.RunAsync();
                await DrawingsChartsAndThemesSample.RunAsync();

                FormulaCalculationSample.Run();
                await TablesPivotTableAndSlicersSample.RunAsync();
                EncryptionProtectionAndVBASample.Run();
            }            
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
			}

            var prevColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"Generated sample workbooks can be found in {FileUtil.OutputDir.FullName}");
            Console.ForegroundColor = prevColor;

            Console.WriteLine();
			Console.WriteLine("Press the return key to exit...");
			
            Console.ReadKey();
		}
	}
}
