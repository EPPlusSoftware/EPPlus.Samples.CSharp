using EPPlusSamples.PivotTables;
using EPPlusSamples.TablesPivotTablesAndSlicers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class TablesPivotTableAndSlicersSample
    {
        public static async Task RunAsync()
        {
            //Sample 7.1 - Custom Named Table, Pivot Table and Slicer styles
            Console.WriteLine("Running sample 7.1 - Working with tables");
            await TablesSample.RunAsync();
            Console.WriteLine("Sorting tables sample...");
            await SortingTablesSample.RunAsync();
            Console.WriteLine("Sample 7.1 finished.");
            Console.WriteLine();

            //Sample 7.2 - Table slicers and Pivot table slicers
            SlicerSample.Run();

            //sample 7.3 - pivot tables
            //This sample demonstrates how to create and work with pivot tables.
            PivotTablesSample.Run();
            //The second class demonstrates how to style you pivot table.
            PivotTablesStylingSample.Run();

            //This sample demonstrates how to calculate and fetch calculated data from a pivot table.
            PivotTablesCalculationSample.Run();

            //Sample 7.4 - Custom Named Table, Pivot Table and Slicer styles
            CustomTableSlicerStyleSample.Run();
        }
    }
}
