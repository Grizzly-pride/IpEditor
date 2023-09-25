namespace IpEditor.Entity.ExcelData.Sheets;

internal sealed class SheetVLANMAP : Sheet
{
    public required string NextHopIP { get; init; }
    public required string Mask { get; init; }
    public required string VLANID { get; init; }
}
