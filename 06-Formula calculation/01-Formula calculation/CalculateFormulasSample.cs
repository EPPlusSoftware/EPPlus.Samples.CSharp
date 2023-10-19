﻿/*************************************************************************************************
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
using EPPlusSamples.FormulaCalculation;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusSamples
{
    /// <summary>
    /// Sample 17 demonstrates the formula calculation engine of EPPlus.
    /// </summary>
    public static class CalculateFormulasSample
    {
        public static void Run()
        {
            CalculateExistingWorkbook.Run();
            BuildAndCalculateWorkbook.Run();
            AddFormulaFunction.Run();
        }

    }
}
