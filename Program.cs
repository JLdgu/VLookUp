using ClosedXML.Excel;

const string vlookup_workbook = @"C:\Users\Jonathan.Linstead\OneDrive - Devon County Council\EUCSharepoint\Mobile Phones\iPad for reuse and disposal.xlsx";
const string vSearch_workbook = @"C:\Users\Jonathan.Linstead\Downloads\CI List2023_08_8_09_16_04.xlsx";

const string lookup_worksheet = "Sheet1";
Range lookup = new(2, 13, 15); // Column to lookup, Start Row, End Row
const int output_column = 12; 

const string search_worksheet = "Data";
Range search = new(11, 56, 58);  // Column to search, Start Row, End Row
const int result_column = 4;
bool deleteUnmatchRows = false;
bool deleteOnly = false;

using XLWorkbook workbook = new(vlookup_workbook);
using XLWorkbook searchBook = new(vSearch_workbook);

IXLWorksheet lookupWorksheet = workbook.Worksheet(lookup_worksheet);
IXLWorksheet searchWorksheet = searchBook.Worksheet(search_worksheet);

Console.WriteLine( "Look Column: {0}",lookupWorksheet.Cell(1, lookup.Column).Value.ToString());
Console.WriteLine( "Search Column: {0}",searchWorksheet.Cell(1, search.Column).Value.ToString());
Console.WriteLine( "Result Column: {0}",searchWorksheet.Cell(1, result_column).Value.ToString());

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
