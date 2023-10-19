using EPPlusSamples.ImportAndExport;
using EPPlusSamples.LoadDataFromCsvFilesIntoTables;
using EPPlusSamples.LoadingData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class ImportAndExportSamples
    {
        public static async Task RunAsync()
        {
            ToCollectionSample.Run();

            LoadingDataFromCollection.Run();

            //Sample 2.3 Loads two csv files into tables and creates an area chart and a Column/Line chart on the data.
            //This sample also shows how to use a secondary axis.
            await ImportAndExportCsvFilesSample.RunAsync();

            //Sample 2.4 - Import and Export DataTable
            DataTableSample.Run();

            // Sample 2.5 - Html Export
            //This sample shows basic html export functionality.
            //For more advanced samples using charts see https://samples.epplussoftware.com/HtmlExport
            HtmlTableExportSample.Run();
            await HtmlRangeExportSample.RunAsync();

            //Sample 32 - Json Export
            //This sample shows the json export functionality.
            //For more a samples exporting to chart librays see https://samples.epplussoftware.com/JsonExport
            await JsonExportSample.RunAsync();

            // Sample 2.7 - ToCollection and ToCollectionWithMappings
            // This sample shows how to export data from a worksheet
            // to a IEnumerable<T> where T is a class.
            ToCollectionSample.Run();
        }
    }
}
