using IpEditor;
using OfficeOpenXml;


string basePath = AppDomain.CurrentDomain.BaseDirectory;
string sourceFilePath = Path.Combine(basePath, "data", "Source.xlsx");
string targetFilePath = Path.Combine(basePath, "data", "Target.xlsx");


ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

List<BaseStation> baseStations = await Editor.LoadSourceData(sourceFilePath);

if(baseStations.Count is not 0)
{
    if (await Editor.OpenTargetFile(targetFilePath))
    {
        await Task.WhenAll(
            Editor.EditIPCLKLNK(baseStations),
            Editor.EditOMCH(baseStations),
            Editor.EditSCTPLNK(baseStations),
            Editor.EditSCTPHOST(baseStations),
            Editor.EditUSERPLANEHOST(baseStations),
            Editor.EditIPPATH(baseStations),
            Editor.EditSRCIPRT(baseStations),
            Editor.EditDEVIP(baseStations));

        Editor.CloseTargetFile();
    }
}

Console.ReadLine();