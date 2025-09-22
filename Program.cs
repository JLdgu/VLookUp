using ClosedXML.Excel;
using Serilog;
using Serilog.Sinks.SystemConsole.Themes;
using System.Text.Json;
using VLookUp;

bool deleteUnmatchRows = false;
bool deleteOnly = false;

Log.Logger = new LoggerConfiguration()
    .Enrich.FromLogContext()
    .WriteTo.Console(theme: AnsiConsoleTheme.Sixteen)
    .WriteTo.File("vlookup.log")
    .MinimumLevel.Debug()
    .CreateLogger();

string fileName = "parameters.json";
string jsonString = File.ReadAllText(fileName);

try
{
    Parameters? parameters = JsonSerializer.Deserialize<Parameters>(jsonString);

    if (parameters == null) return; // Serialize will throw exception if expected parameters are missing

    if (!File.Exists(parameters.LookupWorkbook))
    {
        Log.Error("Lookup workbook {0} not found", parameters.LookupWorkbook);
        return;
    }
    if (!File.Exists(parameters.SearchWorkbook))
    {
        Log.Error("Search workbook {0} not found", parameters.SearchWorkbook);
        return;
    }
    using XLWorkbook workbook = new(parameters.LookupWorkbook);
    using XLWorkbook searchBook = new(parameters.SearchWorkbook);

    IXLWorksheet lookupWorksheet = workbook.Worksheet(parameters.LookupWorksheet);
    if (lookupWorksheet is null)
    {
        Log.Error("Lookup worksheet {0} not found in workbook {1}", parameters.LookupWorksheet, parameters.LookupWorkbook);
        return;
    }
    IXLRow? lookupLastRowUsed = lookupWorksheet.LastRowUsed();
    if (lookupLastRowUsed is null)
    {
        Log.Error("No used rows found in lookup worksheet");
        return;
    }
    if (parameters.LookupStartRow == 0)
        parameters.LookupStartRow = parameters.LookupWorksheetHasHeader ? 2 : 1;

    if (parameters.LookupEndRow == 0)
        parameters.LookupEndRow = lookupLastRowUsed.RowNumber();

    IXLWorksheet searchWorksheet = searchBook.Worksheet(parameters.SearchWorksheet);
    if (searchWorksheet is null)
    {
        Log.Error("Search worksheet {0} not found in workbook {1}", parameters.SearchWorksheet, parameters.SearchWorkbook);
        return;
    }
    IXLRow? searchLastRowUsed = searchWorksheet.LastRowUsed();
    if (searchLastRowUsed is null)
    {
        Log.Error("No used rows found in search worksheet");
        return;
    }
    if (parameters.SearchStartRow == 0)
        parameters.SearchStartRow = parameters.SearchWorksheetHasHeader ? 2 : 1;
    if (parameters.SearchEndRow == 0)
        parameters.SearchEndRow = searchLastRowUsed.RowNumber();

    if (parameters.LookupWorksheetHasHeader)
    {
        Log.Information("Look Column: {0}", lookupWorksheet.Cell(1, parameters.LookupColumn).Value.ToString());
        Log.Information("Output Column: {0}", lookupWorksheet.Cell(1, parameters.LookupOutputColumn).Value.ToString());
    }
    if (parameters.SearchWorksheetHasHeader)
    {
        Log.Information("Search Column: {0}", searchWorksheet.Cell(1, parameters.SearchColumn).Value.ToString());
        Log.Information("Result Column: {0}", searchWorksheet.Cell(1, parameters.SearchResultColumn).Value.ToString());
    }
    Log.Information("Lookup rows {0} to {1}", parameters.LookupStartRow, parameters.LookupEndRow);
    Log.Information("Search rows {0} to {1}", parameters.SearchStartRow, parameters.SearchEndRow);   
    
    for (int row = parameters.LookupEndRow; row > parameters.LookupStartRow - 1; row--)
    {
        string imei = lookupWorksheet.Cell(row, parameters.LookupColumn).Value.ToString();
        string? result = Search(imei);
        if (result is null)
        {
            if (deleteUnmatchRows)
                lookupWorksheet.Row(row).Delete();
        }
        else
        {
            if (!deleteOnly)
                lookupWorksheet.Cell(row, parameters.LookupOutputColumn).Value = result;
        }
    }
    workbook.Save();

    string? Search(string imei)
    {
        for (int row = parameters.SearchStartRow; row < parameters.SearchEndRow + 1; row++)
        {
            if (imei == searchWorksheet.Cell(row, parameters.SearchColumn).Value.ToString())
                return searchWorksheet.Cell(row, parameters.SearchResultColumn).Value.ToString();
        }
        return null;
    }
}
catch (Exception ex)
{
    Log.Fatal(exception: ex, "Unhandled exception:");
    return;
}


record Range(string Column, int StartRow, int EndRow);
