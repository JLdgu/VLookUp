namespace VLookUp;

internal class Parameters
{
    public required string LookupWorkbook { get; set; }
    public required string LookupWorksheet { get; set; }
    public required string LookupColumn { get; set; }
    public required string LookupOutputColumn { get; set; }
    public required int LookupStartRow { get; set; }
    public required int LookupEndRow { get; set; }
    public required string SearchWorkbook { get; set; }
    public required string SearchWorksheet { get; set; }
    public required string SearchColumn { get; set; }
    public required string SearchResultColumn { get; set; }
    public required int SearchStartRow { get; set; }
    public required int SearchEndRow { get; set; }
}
