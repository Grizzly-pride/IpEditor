using IpEditor;
using IpEditor.Entity;
using OfficeOpenXml;
using System.Text.Json;


ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string basePath = AppDomain.CurrentDomain.BaseDirectory;
string targetFilePath = Path.Combine(basePath, "data", "Target.xlsx");
string jsonSettings = Path.Combine(basePath, "Settings.json");

Settings.PrintLogo(ConsoleColor.DarkYellow);
Settings settings;

using (FileStream openStream = File.OpenRead(jsonSettings))
{
    settings = await JsonSerializer.DeserializeAsync<Settings>(openStream)
        ?? throw new FileNotFoundException();
}

List<BaseStation> baseStations = await Editor.LoadSourceData(settings.SourceFile);

if(baseStations.Count is not 0)
{
    var targetFile = settings!.TargetFile;

    if (await Editor.OpenTargetFile(settings?.TargetFile.PathFile ?? targetFilePath))
    {
        var bsNotFoundList = Editor.CheckTargetBS(baseStations, targetFile.SheetBSTransportData);

        if (bsNotFoundList != null)
        {
            var usedBsList = Editor.GetUsedeNodeB(baseStations, bsNotFoundList);
            Editor.EditIPCLKLNK(usedBsList, targetFile.SheetIPCLKLNK);
            Editor.EditOMCH(usedBsList, targetFile.SheetOMCH);
            Editor.EditSCTPLNK(usedBsList, targetFile.SheetSCTPLNK);
            Editor.EditSCTPHOST(usedBsList, targetFile.SheetSCTPHOST);
            Editor.EditUSERPLANEHOST(usedBsList, targetFile.SheetUSERPLANEHOST);
            Editor.EditIPPATH(usedBsList, targetFile.SheetIPPATH);
            Editor.EditSRCIPRT(usedBsList, targetFile.SheetSRCIPRT);
            Editor.EditDEVIP(usedBsList, targetFile.SheetDEVIP);
            Editor.EditVLANMAP(usedBsList, targetFile.SheetVLANMAP);

            Settings.TaskCompletedMessage(ConsoleColor.Blue);
        }
        await Editor.CloseTargetFile();
    }
}

Console.ReadLine();