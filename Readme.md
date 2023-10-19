# EPPlus samples

### EPPlus samples for .Net Core

The solution can be opened in Visual Studio for Windows or MacOS. On other operating systems please use...

```
dotnet restore
dotnet run
```

... to execute the samples.

# Table of Contents
1. [Workbook Worksheet and Ranges](<#workbook-worksheet-and-ranges>)
2. [Import and Export](#import-and-export)
3. [Styling](#styling)
4. [Filters and Validations](#filters-and-validations)
5. [Drawings Charts and Themes](#drawings-charts-and-themes)
6. [Formula Calculation](#formula-calculation)
7. [Tables PivotTables and Slicers](#tables-pivot-tables-and-slicers)
8. [Encryption Protection and VBA](#encryption-protection-and-vba)

### [Workbook Worksheet and Ranges](</01-Workbook Worksheet and Ranges/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|1.01|[Getting started] | Basic usage of EPPlus: create a workbook, fill with data and some basic styling|
|1.02|[Read an Existing Workbook] | Read in a workbook|
|1.03|[Using AsyncAwait] | Async Await in Epplus|
|1.04|[Fill Copy and Sort Ranges] |Different range operations|
|1.05|[Notes and Threaded Comments] | Handling notes and Comments in Epplus|
|1.06|[SalesReport With Hyperlinks] | Hyperlink and database guide|
|1.07|[Performance Write] | Loading and writing in a performant way|
|1.08|[Performance Read Using Linq] | Using Linq on Cells|
|1.09|[ExternalLinks] | External links to other workbooks|
|1.10|[Ignore Errors] | How to ignore errors on ranges|

[Getting started]: </01-Workbook Worksheet and Ranges/01-Create a simple workbook/Readme.md/>
[Read an Existing Workbook]: </01-Workbook Worksheet and Ranges/02-Read an existing workbook/Readme.md/>
[Using AsyncAwait]: </01-Workbook Worksheet and Ranges/03-Using async await/Readme.md/>
[Fill Copy and Sort Ranges]: </01-Workbook Worksheet and Ranges/04-Fill Copy and Sort Ranges/Readme.md/>
[Notes and Threaded Comments]: </01-Workbook Worksheet and Ranges/05-Notes and Threaded Comments/Readme.md/>
[SalesReport With Hyperlinks]: </01-Workbook Worksheet and Ranges/06-SalesReport With Hyperlinks/Readme.md/>
[Performance Write]: </01-Workbook Worksheet and Ranges/07-Performance Write/Readme.md/>
[Performance Read Using Linq]: </01-Workbook Worksheet and Ranges/08-Performance Read Using Linq/Readme.md/>
[ExternalLinks]: </01-Workbook Worksheet and Ranges/09-ExternalLinks/Readme.md/>
[Ignore Errors]: </01-Workbook Worksheet and Ranges/10-IgnoreErrors/Readme.md/>
___

### [Import and Export](</02-Import and Export/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|2.1|[Load Data From Collection]| Load data into worksheet from various types of objects. It also demonstrates the Autofit columns feature.|
|2.2|[Export Data To Collection]| Demonstrates Export ranges and tables into an IEnumerable&lt;T&gt; where T is a class type. |
|2.3|[Import export csv files] | Async Await in Epplus|
|2.4|[Import export DataTable] |Different range operations|
|2.5|[Export to Html] | Handling notes and Comments in Epplus|
|2.6|[Export to Json] | Hyperlink and database guide|

[Import and Export]: </02-Import and Export/Readme.md>
[Load Data From Collection]: </02-Import and Export/01-Load data from collection/Readme.md/>
[Export Data To Collection]: </02-Import and Export/02-Export data to collection/Readme.md/>
[Import export csv files]: </02-Import and Export/03-Import export csv files/Readme.md/>
[Import export DataTable]: </02-Import and Export/04-Import export DataTable/Readme.md/>
[Export to Html]: </02-Import and Export/05-Export to Html/Readme.md/>
[Export to Json]: </02-Import and Export/06-Export to Json/Readme.md/>

___

### [Styling](</03-Styling/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|3.1|[Styling basics] | Demonstrates how to apply different styling on cells including Fills, Fonts etc. |
|3.2|[ConditionalFormatting]| Demonstrates all conditional formatting and how to apply styling/Formatting. |
|3.3|[FXReportFromDatabase] | This sample produces a workbook with foreign exchange rates. |
|3.4|[CreateFileSystemReport] | Demonstrates usage of styling, printer settings, rich text, pie-, doughnut- and bar-charts, freeze panes.|

[Styling basics]: </03-Styling/01-Styling basics/Readme.md/>
[ConditionalFormatting]: </03-Styling/02-ConditionalFormatting/Readme.md/>
[FXReportFromDatabase]: </03-Styling/03-FXReportFromDatabase/Readme.md/>
[CreateFileSystemReport]: </03-Styling/04-CreateFileSystemReport/Readme.md/>

___

### [Filters and Validations](</04-Filters and Validations/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|4.1|[Data Validations] | Demonstrates how to apply different styling on cells including Fills, Fonts etc. |
|4.2|[Filter]| Demonstrates all conditional formatting and how to apply styling/Formatting. |

[Data Validations]: </04-Filters and Validations/01-DataValidation/Readme.md/>
[Filter]: </04-Filters and Validations/02-Filter/Readme.md/>

___

### [Drawings Charts and Themes](</05-Drawings Charts and Themes/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|5.1|[Shapes and Images]| How to create shapes and images using EPPlus and apply effects|
|5.2|[Open Workbook Add Data And Chart]|Open an existing workbook with EPPlus, add data and charts |
|5.3|[Charts And Themes]|Demonstrates how to use various types of charts, chart styling and themes  |
|5.4|[Sparklines]| demonstrates EPPlus support for Sparklines|
|5.5|[FormControls] | demonstrates how to Add form controls, like drop-downs, buttons and radiobuttons to a worksheet and grouping drawings via VBA Macro|


[Shapes and Images]: </05-Drawings Charts and Themes/01-ShapesAndImages/Readme.md/>
[Open Workbook Add Data And Chart]: </05-Drawings Charts and Themes/02-OpenWorkbookAddDataAndChart/Readme.md/>
[Charts And Themes]: </05-Drawings Charts and Themes/03-ChartsAndThemes/Readme.md/>
[Sparklines]: </05-Drawings Charts and Themes/04-Sparklines/Readme.md/>
[FormControls]: </05-Drawings Charts and Themes/05-FormControls/Readme.md/>

___

### [Formula Calculation](</06-Formula Calculation/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|6.1|[Formula Calculation] | shows formula calculation capabilities of EPPlus|
|6.2|[Dynamic Array Formulas] | Example of dynamic array formula calculation |

[Formula Calculation]: </06-Formula Calculation/01-FormulaCalculation/Readme.md/>
[Dynamic Array Formulas]: </06-Formula Calculation/02-DynamicArrayFormulas/Readme.md/>

___

### [Tables Pivot Tables and Slicers](</07-Tables Pivot Tables and Slicers/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|7.1|[Tables]| Samples of creating and calculating tables in Epplus|
|7.2|[Piot Tables]| Pivot table functionality of EPPlus |
|7.3|[Slicers]| Slicers for table and pivot table slicers |
|7.4|[Custom Named Styles]| Example for how to create custom slicer styles |

[Tables]: </07-Tables Pivot Tables and Slicers/01-Tables/Readme.md/>
[Piot Tables]: </07-Tables Pivot Tables and Slicers/02-PivotTables/Readme.md/>
[Slicers]: </07-Tables Pivot Tables and Slicers/03-Slicers/Readme.md/>
[Custom Named Styles]: </07-Tables Pivot Tables and Slicers/04-CustomNamedStyles/Readme.md/>

___

### [Encryption Protection and VBA](</08-Encryption Protection and VBA/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|8.1|[Encryption and Protection]| Sample of encryption and password protection|
|8.2|[VBA]| An implementation of a Battleship game, implemented in Excel/VBA as an example of epplus+VBA interaction |

[Encryption and Protection]: </08-Encryption Protection and VBA/01-EncryptionAndProtection/Readme.md/>
[VBA]: </08-Encryption Protection and VBA/02-VBA/Readme.md/>

### Output files
The samples above produces some workbooks - the name of each workbook indicates which sample that generated it. These workbooks are located in a subdirectory - named "SampleApp" - to the output directory of the sample project.

Also see wiki on https://github.com/EPPlusSoftware/EPPlus/wiki for more details
