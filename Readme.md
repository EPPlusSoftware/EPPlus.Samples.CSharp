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

### [Workbook Worksheet and Ranges](</01-Workbook worksheet and ranges/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|1.01|[Getting started] | Basic usage of EPPlus: create a workbook, fill with data and some basic styling|
|1.02|[Read an existing workbook] | Read in a workbook|
|1.03|[Using async await] | Async Await in Epplus|
|1.04|[Fill copy and sort ranges] |Different range operations|
|1.05|[Notes and threaded comments] | Handling notes and Comments in Epplus|
|1.06|[Sales report with hyperlinks] | Hyperlink and database guide|
|1.07|[Performance write] | Loading and writing in a performant way|
|1.08|[Performance read using linq] | Using Linq on Cells|
|1.09|[External links] | External links to other workbooks|
|1.10|[Ignore errors] | How to ignore errors on ranges|

[Getting started]: </01-Workbook worksheet and ranges/01-Create a simple workbook/Readme.md/>
[Read an existing workbook]: </01-Workbook worksheet and ranges/02-Read an existing workbook/Readme.md/>
[Using async await]: </01-Workbook worksheet and ranges/03-Using async await/Readme.md/>
[Fill copy and sort ranges]: </01-Workbook worksheet and ranges/04-Fill copy and sort ranges/Readme.md/>
[Notes and threaded comments]: </01-Workbook worksheet and ranges/05-Notes and threaded comments/Readme.md/>
[Sales report with hyperlinks]: </01-Workbook worksheet and ranges/06-Sales report with hyperlinks/Readme.md/>
[Performance write]: </01-Workbook worksheet and ranges/07-Performance write/Readme.md/>
[Performance read using linq]: </01-Workbook worksheet and ranges/08-Performance read using Linq/Readme.md/>
[External links]: </01-Workbook worksheet and ranges/09-External links/Readme.md/>
[Ignore errors]: </01-Workbook worksheet and ranges/10-Ignore errors/Readme.md/>
___

### [Import and Export](</02-Import and export/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|2.1|[Load data from collection]| Load data into worksheet from various types of objects. It also demonstrates the Autofit columns feature.|
|2.2|[Export data to collection]| Demonstrates Export ranges and tables into an IEnumerable&lt;T&gt; where T is a class type. |
|2.3|[Import export csv files] | Async Await in Epplus|
|2.4|[Import export DataTable] |Different range operations|
|2.5|[Export to html] | Handling notes and Comments in Epplus|
|2.6|[Export to json] | Hyperlink and database guide|

[Import and export]: </02-Import and export/Readme.md>
[Load data from collection]: </02-Import and export/01-Load data from collection/Readme.md/>
[Export data to collection]: </02-Import and export/02-Export data to collection/Readme.md/>
[Import export csv files]: </02-Import and export/03-Import export csv files/Readme.md/>
[Import export DataTable]: </02-Import and export/04-Import export DataTable/Readme.md/>
[Export to html]: </02-Import and export/05-Export to html/Readme.md/>
[Export to json]: </02-Import and export/06-Export to json/Readme.md/>

___

### [Styling](</03-Styling/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|3.1|[Styling basics] | Demonstrates how to apply different styling on cells including Fills, Fonts etc. |
|3.2|[Conditional formatting]| Demonstrates all conditional formatting and how to apply styling/Formatting. |
|3.3|[Fx report from database] | This sample produces a workbook with foreign exchange rates. |
|3.4|[Create a file system report] | Demonstrates usage of styling, printer settings, rich text, pie-, doughnut- and bar-charts, freeze panes.|

[Styling basics]: </03-Styling/01-Styling basics/Readme.md/>
[Conditional formatting]: </03-Styling/02-Conditional formatting/Readme.md/>
[Fx report from database]: </03-Styling/03-Fx report from database/Readme.md/>
[Create a file system report]: </03-Styling/04-Create a file system report/Readme.md/>

___

### [Filters and Validations](</04-Filters and validations/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|4.1|[Data Validations] | Demonstrates how to apply different styling on cells including Fills, Fonts etc. |
|4.2|[Filter]| Demonstrates all conditional formatting and how to apply styling/Formatting. |

[Data Validations]: </04-Filters and validations/01-Data validation/Readme.md/>
[Filter]: </04-Filters and validations/02-Filter/Readme.md/>

___

### [Drawings Charts and Themes](</05-Drawings charts and themes/Readme.md>)
|No|Sample|Description|
|---|---|-----------------|
|5.1|[Shapes and images]| How to create shapes and images using EPPlus and apply effects|
|5.2|[Open workbook add data and chart]|Open an existing workbook with EPPlus, add data and charts |
|5.3|[Charts And themes]|Demonstrates how to use various types of charts, chart styling and themes  |
|5.4|[Sparklines]| demonstrates EPPlus support for Sparklines|
|5.5|[Form controls] | demonstrates how to Add form controls, like drop-downs, buttons and radiobuttons to a worksheet and grouping drawings via VBA Macro|


[Shapes and images]: </05-Drawings charts and themes/01-Shapes and images/Readme.md/>
[Open workbook add data and chart]: </05-Drawings charts and themes/02-Open workbook add data and chart/Readme.md/>
[Charts And themes]: </05-Drawings charts and themes/03-Charts and themes/Readme.md/>
[Sparklines]: </05-Drawings charts and themes/04-Sparklines/Readme.md/>
[Form controls]: </05-Drawings charts and themes/05-Form controls/Readme.md/>

___

### [Formula Calculation](</06-Formula calculation/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|6.1|[Formula Calculation] | shows formula calculation capabilities of EPPlus|
|6.2|[Dynamic Array Formulas] | Example of dynamic array formula calculation |

[Formula Calculation]: </06-Formula calculation/01-Formula calculation/Readme.md/>
[Dynamic Array Formulas]: </06-Formula calculation/02-Array formulas/Readme.md/>

___

### [Tables Pivot Tables and Slicers](</07-Tables pivot tables and slicers/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|7.1|[Tables]| Samples of creating and calculating tables in Epplus|
|7.2|[Pivot Tables]| Pivot table functionality of EPPlus |
|7.3|[Slicers]| Slicers for table and pivot table slicers |
|7.4|[Custom Named Styles]| Example for how to create custom slicer styles |

[Tables]: </07-Tables pivot tables and slicers/01-Tables/Readme.md/>
[Pivot Tables]: </07-Tables pivot tables and slicers/02-Pivot tables/Readme.md/>
[Slicers]: </07-Tables pivot tables and slicers/03-Slicers/Readme.md/>
[Custom Named Styles]: </07-Tables pivot tables and slicers/04-Custom named styles/Readme.md/>

___

### [Encryption Protection and VBA](</08-Encryption protection and VBA/Readme.md>)

|No|Sample|Description|
|---|---|-----------------|
|8.1|[Encryption and Protection]| Sample of encryption and password protection|
|8.2|[VBA]| An implementation of a Battleship game, implemented in Excel/VBA as an example of epplus+VBA interaction |

[Encryption and Protection]: </08-Encryption protection and VBA/01-Encryption and protection/Readme.md/>
[VBA]: </08-Encryption protection and VBA/02-VBA/Readme.md/>

### Output files
The samples above produces some workbooks - the name of each workbook indicates which sample that generated it. These workbooks are located in a subdirectory - named "SampleApp" - to the output directory of the sample project.

Also see wiki on https://github.com/EPPlusSoftware/EPPlus/wiki for more details
