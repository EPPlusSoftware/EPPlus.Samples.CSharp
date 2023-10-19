using EPPlusSamples.ConditionalFormatting;
using EPPlusSamples.Styling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class StylingBasics
    {
        public static void Run()
        {
            //Sample 2.1 - Basic styling            
            BasicStyleSample.Run();

            // Sample 2.2 - Conditional Formatting
            ConditionalFormattingSamples.Run();

            // Sample 2.3 - Creates a workbook based on a template.
            // Populates a range with data and set the series of a linechart.
            FxReportFromDatabase.Run();

            //Sample 2.4
            //Creates an advanced report on a directory in the filesystem.
            //Parameter 2 is the directory to report. Parameter 3 is how deep the scan will go. Parameter 4 Skips Icons if set to true (The icon handling is slow)
            //This example demonstrates how to use outlines, tables,comments, shapes, pictures and charts.                
            
            var directoryToList = new DirectoryInfo(System.Reflection.Assembly.GetEntryAssembly().Location).Parent;
            CreateAFileSystemReport.Run(directoryToList, 5, true);
        }
    }
}
