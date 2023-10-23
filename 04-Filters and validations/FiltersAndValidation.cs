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
            //Sample 4.1 - Data validation
            DataValidationSample.Run();

            //Sample 4.2 - Filter
            Console.WriteLine("Running sample 4.2-Filter");
            await Filter.RunAsync();
            Console.WriteLine("Sample 4.2 created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();
        }
    }
}
