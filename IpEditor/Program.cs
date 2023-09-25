using IpEditor;
using IpEditor.Entity;
using OfficeOpenXml;
using System.Text.Json;


ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string basePath = AppDomain.CurrentDomain.BaseDirectory;
string sourceFilePath = Path.Combine(basePath, "data", "Source.xlsx");
string targetFilePath = Path.Combine(basePath, "data", "Target.xlsx");
string jsonSettings = Path.Combine(basePath, "Settings.json");

Settings.PrintLogo(ConsoleColor.DarkYellow);
Settings settings;

using (FileStream openStream = File.OpenRead(jsonSettings))
{
    settings = await JsonSerializer.DeserializeAsync<Settings>(openStream)
        ?? throw new FileNotFoundException();
}

List<BaseStation> baseStations = await Editor.LoadSourceData(settings?.SourceFile.PathFile ?? sourceFilePath);

if(baseStations.Count is not 0)
{
    if (await Editor.OpenTargetFile(settings?.TargetFile.PathFile ?? targetFilePath))
    {
        await Task.WhenAll(
            Editor.EditIPCLKLNK(baseStations, settings!.TargetFile.SheetIPCLKLNK.Bs!),
            Editor.EditOMCH(baseStations, settings.TargetFile.SheetOMCH.Bs!),
            Editor.EditSCTPLNK(baseStations, settings.TargetFile.SheetSCTPLNK.Bs!),
            Editor.EditSCTPHOST(baseStations, settings.TargetFile.SheetSCTPHOST.Bs!),
            Editor.EditUSERPLANEHOST(baseStations, settings.TargetFile.SheetUSERPLANEHOST.Bs!),
            Editor.EditIPPATH(baseStations, settings.TargetFile.SheetIPPATH.Bs!),
            Editor.EditSRCIPRT(baseStations, settings.TargetFile.SheetSRCIPRT.Bs!),
            Editor.EditDEVIP(baseStations));

        Editor.CloseTargetFile();
    }
}

Console.ReadLine();