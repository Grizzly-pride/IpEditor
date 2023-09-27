using IpEditor.Entity.ExcelData.Sheets;
namespace IpEditor.Entity.ExcelData;

internal sealed class TargetFile
{
    public string? PathFile { get; init; }

    #region Sheets
    public required SheetBSTransportData SheetBSTransportData { get; init; }
    public required SheetIPCLKLNK SheetIPCLKLNK { get; init; }
    public required SheetVLANMAP SheetVLANMAP { get; init; }
    public required SheetSRCIPRT SheetSRCIPRT { get; init; }
    public required SheetDEVIP SheetDEVIP { get; init; }
    public required SheetIPPATH SheetIPPATH { get; init; }
    public required SheetOMCH SheetOMCH { get; init; }
    public required SheetSCTPLNK SheetSCTPLNK { get; init; }
    public required SheetSCTPHOST SheetSCTPHOST { get; init; }
    public required SheetUSERPLANEHOST SheetUSERPLANEHOST { get; set; }
    #endregion
}
