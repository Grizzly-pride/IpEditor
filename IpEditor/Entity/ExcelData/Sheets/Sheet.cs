namespace IpEditor.Entity.ExcelData.Sheets;

internal abstract class Sheet
{
    public string? SheetName { get; init; }
    public string? Operation { get; init; }
    public string? Bs { get; init; }   
}
