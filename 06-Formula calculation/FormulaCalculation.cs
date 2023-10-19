using EPPlusSamples.FormulaCalculation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class FormulaCalculationSample
    {
        public static void Run()
        {
            //Sample 6 Calculate - Shows how to calculate formulas in the workbook.
            Console.WriteLine("Sample 6.1 - Calculate formulas");
            CalculateFormulasSample.Run();
            Console.WriteLine("Sample 6.1 created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();

            Console.WriteLine("Sample 6.2 - Dynamic array formulas");
            UsingArrayformulas.Run();
            DynamicArrayFromTableWithChart.Run();
            Console.WriteLine("Sample 6.2 created {0}", FileUtil.OutputDir.Name);
            Console.WriteLine();
        }
    }
}
    