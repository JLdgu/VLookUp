using ClosedXML.Excel;
using System.Text.Json;
using VLookUp;

bool deleteUnmatchRows = false;
bool deleteOnly = false;

string fileName = "parameters.json";
string jsonString = File.ReadAllText(fileName);

try
{
    Parameters? parameters = JsonSerializer.Deserialize<Parameters>(jsonString);

    if (parameters == null) return; // Serialize will throw exception if expected parameters are missing

    using XLWorkbook workbook = new(parameters.LookupWorkbook);
    using XLWorkbook searchBook = new(parameters.SearchWorkbook);

    IXLWorksheet lookupWorksheet = workbook.Worksheet(parameters.LookupWorksheet);
    IXLWorksheet searchWorksheet = searchBook.Worksheet(parameters.SearchWorksheet);

    Console.WriteLine("Search spreadsheet last row used {0}", lookupWorksheet.LastRowUsed());

    Console.WriteLine("Look Column: {0}", lookupWorksheet.Cell(1, parameters.LookupColumn).Value.ToString());
    Console.WriteLine("Output Column: {0}", lookupWorksheet.Cell(1, parameters.LookupOutputColumn).Value.ToString());
    Console.WriteLine("Search Column: {0}", searchWorksheet.Cell(1, parameters.SearchColumn).Value.ToString());
    Console.WriteLine("Result Column: {0}", searchWorksheet.Cell(1, parameters.SearchResultColumn).Value.ToString());

    Console.WriteLine("Are these the correct columns? (y/n)");
    var yn = Console.ReadKey();

    if (yn.Key.ToString() == "n" || yn.Key.ToString() == "N" || yn.Key == ConsoleKey.Escape)
        return;

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
catch (Exception)
{
    Console.WriteLine("Unable to parse all Json parameters");
    return;
}


record Range(string Column, int StartRow, int EndRow);
