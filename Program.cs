using ClosedXML.Excel;

const string vlookup_workbook = @"C:\Users\Jonathan.Linstead\OneDrive - Devon County Council\Phones\Mobile phone export excluding status of Disposed 050623.xlsx";

const string lookup_worksheet = "SR204153";
Range lookup = new(2, 2, 944); // 944
const int output_column = 6; 

const string search_worksheet = "Data";
Range search = new(1, 2, 11705);  // Column to search, Start Row, End
const int result_column = 3;
bool deleteMissingRows = true;

using XLWorkbook workbook = new(vlookup_workbook);

IXLWorksheet lookupWorksheet = workbook.Worksheet(lookup_worksheet);
IXLWorksheet searchWorksheet = workbook.Worksheet(search_worksheet);

for (int row = lookup.EndRow; row > lookup.StartRow - 1; row--)
{
    string imei = lookupWorksheet.Cell(row, lookup.Column).Value.ToString();
    string? result = Search(imei);
    if (result is null)
    {
        if (deleteMissingRows)
            lookupWorksheet.Row(row).Delete();
    }
    else
        lookupWorksheet.Cell(row, output_column).Value = result;
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

public record Range(int Column, int StartRow, int EndRow);