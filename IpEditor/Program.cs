using IpEditor;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var sourceFilePath = @"";
var targetFilePatch = @"";


List<BaseStation> baseStations = await Editor.LoadSourceData(sourceFilePath);

await Task.WhenAll(
    Editor.OpenTargetFile(targetFilePatch),
    Editor.EditIPCLKLNK(baseStations),
    Editor.EditOMCH(baseStations),
    Editor.EditSCTPLNK(baseStations),
    Editor.EditSCTPHOST(baseStations),
    Editor.EditUSERPLANEHOST(baseStations),
    Editor.EditIPPATH(baseStations),
    Editor.EditSRCIPRT(baseStations),
    Editor.EditDEVIP(baseStations));

Editor.CloseTargetFile();