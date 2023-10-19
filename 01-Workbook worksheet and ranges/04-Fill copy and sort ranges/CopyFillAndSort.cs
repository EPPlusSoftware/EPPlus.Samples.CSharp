using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.WorkbookWorksheetAndRanges
{
    public static class CopyFillAndSort
    {
        public static void Run()
        {
            Console.WriteLine("Running sample 1.4 - Copy, Fill and Sort Ranges");
            CopyRangeSample.Run();
            FillRangeSample.Run();
            SortingRangesSample.Run();
            Console.WriteLine("Sample 1.4 finished.");
            Console.WriteLine();
        }
    }
}
