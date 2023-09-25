namespace IpEditor.Entity.ExcelData.Sheets;

internal sealed class SheetVLANMAP : Sheet
{
    public string? NextHopIP { get; init; }
    public string? Mask { get; init; }
    public string? VLANID { get; init; }
}
