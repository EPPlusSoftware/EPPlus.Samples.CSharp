using EPPlusSamples.FiltersAndValidations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class FiltersAndValidation
    {
        public static async Task RunAsync()
        {
            //Sample 12 - Data validation
            DataValidationSample.Run();

            //Sample 13 - Filter
            Console.WriteLine("Running sample 13-Filter");
            await Filter.RunAsync();
            Console.WriteLine("Sample 13 created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();
        }
    }
}
