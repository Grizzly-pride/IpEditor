namespace IpEditor.Entity.ExcelData.Sheets;

internal sealed class SheetSRCIPRT : Sheet
{
    public string? SourceIPAddress { get; init; }
    public string? NextHopIP { get; init; }
    public string? UserLabel { get; init; }
}
