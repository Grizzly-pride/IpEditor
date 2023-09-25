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
        await Task.WhenAll(
            Editor.EditIPCLKLNK(baseStations, targetFile.SheetIPCLKLNK),
            Editor.EditOMCH(baseStations, targetFile.SheetOMCH),
            Editor.EditSCTPLNK(baseStations, targetFile.SheetSCTPLNK),
            Editor.EditSCTPHOST(baseStations, targetFile.SheetSCTPHOST),
            Editor.EditUSERPLANEHOST(baseStations, targetFile.SheetUSERPLANEHOST),
            Editor.EditIPPATH(baseStations, targetFile.SheetIPPATH),
            Editor.EditSRCIPRT(baseStations, targetFile.SheetSRCIPRT),
            Editor.EditDEVIP(baseStations, targetFile.SheetDEVIP),
            Editor.EditVLANMAP(baseStations, targetFile.SheetVLANMAP));

        Editor.CloseTargetFile();

        Settings.TaskCompletedMessage(ConsoleColor.Blue);
    }
}

Console.ReadLine();