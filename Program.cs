using ClosedXML.Excel;

const string vlookup_workbook = @"C:\temp\disposals.xlsx";
const string vSearch_workbook = @"C:\temp\CI List2023_10_10_13_14_47.xlsx";

const string lookup_worksheet = "disposals";
Range lookup = new(1, 2, 39); // Column to lookup (A = 1), Start Row, End Row
const int output_column = 10; 

const string search_worksheet = "Data";
Range search = new(4, 2, 4650);  // Column to search, Start Row, End Row
const int result_column = 2;
bool deleteUnmatchRows = false;
bool deleteOnly = false;

using XLWorkbook workbook = new(vlookup_workbook);
using XLWorkbook searchBook = new(vSearch_workbook);

IXLWorksheet lookupWorksheet = workbook.Worksheet(lookup_worksheet);
IXLWorksheet searchWorksheet = searchBook.Worksheet(search_worksheet);

Console.WriteLine( "Look Column: {0}",lookupWorksheet.Cell(1, lookup.Column).Value.ToString());
Console.WriteLine( "Search Column: {0}",searchWorksheet.Cell(1, search.Column).Value.ToString());
Console.WriteLine( "Result Column: {0}",searchWorksheet.Cell(1, result_column).Value.ToString());

Console.WriteLine("Are these the correct columns? (y/n)");
var yn = Console.ReadKey();

if (yn.Key.ToString()  == "n" || yn.Key.ToString() == "N" || yn.Key == ConsoleKey.Escape)        
    return;

for (int row = lookup.EndRow; row > lookup.StartRow - 1; row--)
{
    string imei = lookupWorksheet.Cell(row, lookup.Column).Value.ToString();
    string? result = Search(imei);
    if (result is null)
    {
        if (deleteUnmatchRows)
            lookupWorksheet.Row(row).Delete();
    }
    else
    {
        if (!deleteOnly)
            lookupWorksheet.Cell(row, output_column).Value = result;
    }
}
workbook.Save();


string? Search(string imei)
{
    for (int row = search.StartRow; row < search.EndRow + 1; row++)
    {   
        if (imei == searchWorksheet.Cell(row, search.Column).Value.ToString())
            return searchWorksheet.Cell(row, result_column).Value.ToString();
    }
    return null;
}

record Range(int Column, int StartRow, int EndRow);
