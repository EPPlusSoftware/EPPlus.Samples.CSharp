using EPPlusSamples.LoadingData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples.ImportAndExport
{
    public static class LoadingDataFromCollection
    {
        public static void Run()
        {
            //Sample 4 - Shows a few ways to load data (Datatable, IEnumerable and more).
            Console.WriteLine("Running sample 4 - LoadingDataWithTables");
            LoadingDataWithTablesSample.Run();
            Console.WriteLine("Sample 2.1 (LoadingDataWithTables) created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();

            //Sample 4 - Shows how to load dynamic/ExpandoObject 
            LoadingDataWithDynamicObjects.Run();
            Console.WriteLine("Sample 2.1 (LoadingDataWithDynamicObjects) created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();

            // Sample 4 - LoadFromCollectionWithAttributes
            LoadingDataFromCollectionWithAttributes.Run();
            Console.WriteLine("Sample 2.1 (LoadingDataFromCollectionWithAttributes) created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();
        }
    }
}
