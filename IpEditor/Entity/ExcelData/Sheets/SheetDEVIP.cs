namespace IpEditor.Entity.ExcelData.Sheets;

internal sealed class SheetDEVIP : Sheet
{
    public required string IPAddress { get; init; }
    public required string Mask { get; init; }
    public required string UserLabel { get; init; }
}
