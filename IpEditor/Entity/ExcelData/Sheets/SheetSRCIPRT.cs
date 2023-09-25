namespace IpEditor.Entity.ExcelData.Sheets;

internal sealed class SheetSRCIPRT : Sheet
{
    public required string SourceIPAddress { get; init; }
    public required string NextHopIP { get; init; }
    public required string UserLabel { get; init; }
}
