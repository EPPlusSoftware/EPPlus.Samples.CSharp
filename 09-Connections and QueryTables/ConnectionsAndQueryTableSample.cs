using OfficeOpenXml;
using OfficeOpenXml.Data.Connection;
using OfficeOpenXml.Table;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.ComponentModel.DataAnnotations;
namespace EPPlusSamples
{
    /// <summary>
    /// A sample showing how to add different types of connections.
    /// </summary>
    public static class ConnectionsAndQueryTableSample
    {
        /// <summary>
        /// With EPPlus you can add connections to other data sources. EPPlus will not execute the connections, but you can add then and let excel load the data when the workbook is opened.
        /// EPPlus can add, modify and remove all types of connections, including more modern power query connections. However, EPPlus currently does not support building or modifying data models in Excel.
        /// You can use connections with query tables or as sources to pivot tables. Legacy query tables are added directly to the worksheet, while new query tables are added add as a table using the worksheet.Tables.AddQueryTable method.
        /// </summary>
        public static void Run()
        {
            using var p = new ExcelPackage();            

            /* Create a connection to a text file and load the connection data into the worksheet. This type of connection is considerered legacy in newer versions of Excel. */
            CreateTextConnection(p);
            //Create an OLEDB connection using the Microsoft.ACE provider agains a text file. Adds a query table to the worksheet using the Tables collection.
            CreateOleDbConnection(p);
            //Creates two power query connection, one reading a html table and one reading a text file.
            CreatePowerQueryConnections(p);
            //Creates a power query connection used as source to a pivot table.
            CreatePivotTableWithConnection(p);

            var fi = FileUtil.GetCleanFileInfo("9.1-ConnectionsAndQueryTables.xlsx");
            p.SaveAs(fi);
        }

        private static void CreatePivotTableWithConnection(ExcelPackage p)
        {
            var csvDir = FileUtil.GetSubDirectory("09-Connections and QueryTables", "");
            var connectionString = $"provider=Microsoft.ACE.OLEDB.12.0;data source={csvDir.FullName}\\;extended properties=\"text;HDR=Yes;FMT=Delimited\"";

            //In this sample we create an oledb connection and use it as source for our pivot table. EPPlus will relay on Excel to execute the query.
            var connection = p.Workbook.Connections.AddDatabase("OleDbConnection5", connectionString);
            connection.DatabaseProperties.CommandType = eCommandType.SqlStatement;
            connection.DatabaseProperties.Command = "select * from [Sample9-2.txt]";

            var ws = p.Workbook.Worksheets.Add("PivotTableWithConnection");
            //As EPPlus does not execute the connection/query, you have to specify the fields for the query. These fields must match the output of the query.
            var pt = ws.PivotTables.Add(ws.Cells["A3"], connection, "PivotTable1", ["Name", "Date", "Amount", "Percent", "Category"]);
            var rf = pt.RowFields.Add(pt.Fields["Date"]);            
            rf.Sort = eSortType.Ascending;
            pt.DataFields.Add(pt.Fields["Amount"]);
        }

        private static void CreatePowerQueryConnections(ExcelPackage p)
        {
            CreatePowerQueryHtmlTableConnection(p);
            CreatePowerQueryTextFileConnection(p);
        }

        private static void CreatePowerQueryHtmlTableConnection(ExcelPackage p)
        {
            //The power query OleDB provider, Microsoft.Mashup.OleDb.1, is used by Excel to execute power query formulas written in the language M.
            //To use this provider you also need to create the PowerQuerySettings and set the M formula for your connection.
            var connectionString = "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=\"Table 1\";Extended Properties=\"\"";

            //The power query formula is written the M language. Please see https://learn.microsoft.com/en-us/powerquery-m/power-query-m-language-specification.
            //The easies way to retrieve the formula is to create the power query connection in Excel and get the formula by opening the workbook in EPPlus and get the formula from the PowerQuerySettings.Formula property.
            //Please note the EPPlus does not validate this formula and that it should not contain the Section1 declaration at the beginning.
            var mFormula = "shared #\"Table 1\" = let\r\n    Source = Web.BrowserContents(\"https://github.com/EPPlusSoftware/EPPlus/wiki/Formatting-and-styling\"),\r\n    #\"Extracted Table From Html\" = Html.Table(Source, {{\"Column1\", \"TABLE > * > TR > :nth-child(1)\"}, {\"Column2\", \"TABLE > * > TR > :nth-child(2)\"}}, [RowSelector=\"TABLE > * > TR\"]),\r\n    #\"Promoted Headers\" = Table.PromoteHeaders(#\"Extracted Table From Html\", [PromoteAllScalars=true]),\r\n    #\"Changed Type\" = Table.TransformColumnTypes(#\"Promoted Headers\",{{\"Id\", Int64.Type}, {\"Format\", type text}})\r\nin\r\n    #\"Changed Type\";";
            var dbConn = p.Workbook.Connections.AddPowerQuery("PowerQueryDbHtmlConnection", connectionString, mFormula);

            //The table name in the query should match the declaration in the M formula.
            dbConn.DatabaseProperties.Command = "SELECT * FROM [Table 1]";

            // EPPlus creates a standard meta data XML document for the formulas of the power query, by default. 
            // However, in some cases, you may need or want to add your own meta data. 
            // You can do this by updating the MetadataItems collection or load metadata Xml directly using the LoadMetadataXml method.
            // For details, see MS-QDEFF 2.5.1: https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/aef664f7-e00b-4683-9724-0dec509dc658
            // See the commented-out code example below.

            //var metadataXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?><LocalPackageMetadataFile xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><Items><Item><ItemPath><ItemType>AllFormulas</ItemType><ItemPath /></ItemPath><StableEntries><Entry Type=\"Relationships\" Value=\"sAAAAAA==\" /></StableEntries></Item><Item><ItemPath><ItemType>Formula</ItemType><ItemPath>Section1/Table%201</ItemPath></ItemPath><StableEntries><Entry Type=\"QueryID\" Value=\"sfa48b4ec-74b8-4199-b152-c11e9a218edb\" /><Entry Type=\"FillEnabled\" Value=\"l1\" /><Entry Type=\"FillObjectType\" Value=\"sTable\" /><Entry Type=\"FillToDataModelEnabled\" Value=\"l0\" /><Entry Type=\"IsPrivate\" Value=\"l0\" /><Entry Type=\"BufferNextRefresh\" Value=\"l1\" /><Entry Type=\"ResultType\" Value=\"sTable\" /><Entry Type=\"NameUpdatedAfterFill\" Value=\"l0\" /><Entry Type=\"FillTarget\" Value=\"sTable_1\" /><Entry Type=\"FilledCompleteResultToWorksheet\" Value=\"l1\" /><Entry Type=\"AddedToDataModel\" Value=\"l0\" /><Entry Type=\"FillCount\" Value=\"l28\" /><Entry Type=\"FillErrorCode\" Value=\"sUnknown\" /><Entry Type=\"FillErrorCount\" Value=\"l0\" /><Entry Type=\"FillLastUpdated\" Value=\"d2025-11-12T12:40:50.0911073Z\" /><Entry Type=\"FillColumnTypes\" Value=\"sAwY=\" /><Entry Type=\"FillColumnNames\" Value=\"s[&quot;Id&quot;,&quot;Format&quot;]\" /><Entry Type=\"FillStatus\" Value=\"sComplete\" /><Entry Type=\"RelationshipInfoContainer\" Value=\"s{&quot;columnCount&quot;:2,&quot;keyColumnNames&quot;:[],&quot;queryRelationships&quot;:[],&quot;columnIdentities&quot;:[&quot;Section1/Table 1/AutoRemovedColumns1.{Id,0}&quot;,&quot;Section1/Table 1/AutoRemovedColumns1.{Format,1}&quot;],&quot;ColumnCount&quot;:2,&quot;KeyColumnNames&quot;:[],&quot;ColumnIdentities&quot;:[&quot;Section1/Table 1/AutoRemovedColumns1.{Id,0}&quot;,&quot;Section1/Table 1/AutoRemovedColumns1.{Format,1}&quot;],&quot;RelationshipInfo&quot;:[]}\" /></StableEntries></Item><Item><ItemPath><ItemType>Formula</ItemType><ItemPath>Section1/Table%201/Source</ItemPath></ItemPath><StableEntries /></Item><Item><ItemPath><ItemType>Formula</ItemType><ItemPath>Section1/Table%201/Extracted%20Table%20From%20Html</ItemPath></ItemPath><StableEntries /></Item><Item><ItemPath><ItemType>Formula</ItemType><ItemPath>Section1/Table%201/Promoted%20Headers</ItemPath></ItemPath><StableEntries /></Item><Item><ItemPath><ItemType>Formula</ItemType><ItemPath>Section1/Table%201/Changed%20Type</ItemPath></ItemPath><StableEntries /></Item></Items></LocalPackageMetadataFile>";
            //p.Workbook.PowerQuerySettings.LoadMetadataXml(metadataXml);

            var ws = p.Workbook.Worksheets.Add("PowerQueryWeb");
            //Add a query table with three columns. The columns must be specified in the last string array parameter and must match the query.
            //More columns with formulas can be added. These columns must set "DataBoundColumn" to false.
            //In this table we have two column from the html table and one additional calculated formula...
            var tbl = ws.Tables.AddQueryTable(ws.Cells["A1:C2"], "Table_1", dbConn, ["Id", "Format", "Formula"]);
            tbl.QueryTable.Fields[2].DataBoundColumn = false;
            tbl.Columns[2].CalculatedColumnFormula = "Table_1[[#This Row],[Id]]+1";

            //EPPlus does not execute connections/queries, so we refresh the query when the workbook is loaded.
            tbl.QueryTable.RefreshOnLoad = true;
        }
        private static void CreatePowerQueryTextFileConnection(ExcelPackage p)
        {
            //Setup the connectionstring for power query. It uses the Microsoft.Mashup.OleDb.1 provider. the location property should contain the object in the M-formula.
            var connectionString = "Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=\"Table 2\";Extended Properties=\"\"";

            var csvFile = FileUtil.GetFileInfo("09-Connections and QueryTables", "Sample9-1.txt");
            //The M formula. EPPlus does not validate this formulas, so make sure you set it up correctly. A good way of getting your formulas is to create the query in Excel and then open the file with EPPlus to extract the formula from the Workbook.PowerQuerySettings.Formulas property.
            //Please note that the M-Formulas supplied to The AddPowerQuery method should not contain the Section1 declaration.
            var mFormula = "shared #\"Table 2\" = let\r\n    Source = Csv.Document(File.Contents(\"" + csvFile + "\"),[Delimiter=\",\", Columns=7, QuoteStyle=QuoteStyle.None]),\r\n    #\"Promoted Headers\" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),\r\n    #\"Changed Type\" = Table.TransformColumnTypes(#\"Promoted Headers\",{{\"Period\", type date}, {\"Europe\", type number}, {\"Africa\", type number}, {\"Asia\", type number}, {\"North America\", type number}, {\"South America\", type number}, {\"Austraila\", type number}}, \"en-US\")\r\nin\r\n    #\"Changed Type\";";

            var dbConn = p.Workbook.Connections.AddPowerQuery("PowerQueryTextConnection", connectionString, mFormula);
            dbConn.DatabaseProperties.Command = "SELECT * FROM [Table 2]";

            var ws = p.Workbook.Worksheets.Add("PowerQueryText");
            //Add a query table with seven columns. As EPPlus does not execute connections/queries, the columns must be specified in the last string array parameter and must match the query output.
            var tbl = ws.Tables.AddQueryTable(ws.Cells["A1:G2"], "Table_2", dbConn, ["Period", "Europe", "Africa", "Asia", "North America", "South America", "Austraila"]);
            tbl.TableStyle = TableStyles.Dark3;

            //EPPlus does not execute connections/queries, so we refresh the query when the workbook is loaded.
            tbl.QueryTable.RefreshOnLoad = true;

            //Styling for a table column can be set via the DataStyle property like this:
            tbl.Columns[0].DataStyle.NumberFormat.Format = "yyyy-MM";

            //The formula and other settings for power query can be found on the Workbooks PowerQuerySettings object.

            Console.WriteLine("The formula and other settings for power query can be found on the Workbooks PowerQuerySettings object:");
            Console.WriteLine("Power Query Formulas: " + p.Workbook.PowerQuerySettings.Formulas);
        }

        private static void CreateOleDbConnection(ExcelPackage p)
        {
            var csvDir = FileUtil.GetSubDirectory("09-Connections and QueryTables","");
            //This sample uses the Microsoft.ACE provider as a sample, but a more common scenario would be to have a connection agains a SQL or OLAP server.
            var connectionString = $"provider=Microsoft.ACE.OLEDB.12.0;data source={csvDir.FullName}\\;extended properties=\"text;HDR=Yes;FMT=Delimited\"";
            var c = p.Workbook.Connections.AddDatabase("OleDbConnection1", connectionString);
            c.DatabaseProperties.CommandType = eCommandType.SqlStatement;
            c.DatabaseProperties.Command = "select * from [Sample9-1.txt]";

            var ws = p.Workbook.Worksheets.Add("OleDbConnection");

            var qt = ws.Tables.AddQueryTable(ws.Cells["A1:G2"], "MyOleDbQuery", c, ["Period", "Europe", "Africa", "Asia", "North America", "South America", "Austraila"]);

            //EPPlus does not execute connections/queries, so we refresh the query when the workbook is loaded.
            qt.QueryTable.RefreshOnLoad = true;
        }
        private static void CreateTextConnection(ExcelPackage p)
        {
            var csvFile = FileUtil.GetFileInfo("09-Connections and QueryTables", "Sample9-1.txt");
            //EPPlus supports adding older types of connection, for example a connection directly agains a text file.
            var c = p.Workbook.Connections.AddText("TextConnection1", csvFile);
            c.TextProperties.Delimited = true;
            c.TextProperties.Delimiter = ",";
            c.TextProperties.Decimal = ".";
            c.TextProperties.Fields.Add(new ExcelConnectionTextField(eConnectionTextFieldType.YearMonthDay));
            c.TextProperties.Prompt = false; //Don't prompt for the file loction on refresh.

            var ws = p.Workbook.Worksheets.Add("TextConnection");

            //Legacy text connections must be added directly to the worksheet, as it is not supported using tables.
            var qt = ws.QueryTables.Add(ws.Cells["A1:G5"], "MyQuertyTable", c);
            ws.Cells["A1:A5"].Style.Numberformat.Format = "yyyy-MM";
            ws.Cells["A1:G1"].Style.Font.Bold = true;

            //EPPlus does not execute connections/queries, so we refresh the query when the workbook is loaded.
            qt.RefreshOnLoad = true;
        }
    }   
}
