using IpEditor;
using IpEditor.Entity;
using IpEditor.Entity.ExcelData.Sheets;
using OfficeOpenXml;
using System.Text.Json;
using System.Threading.Tasks;


//Settings.PrintLogo(ConsoleColor.Yellow);

//Settings settings = new Settings()
//{
//    SourceFile = new()
//    {
//        PathFile = "test",
//        BSName = "test",
//        OAM = new Route()
//        {
//            SourceIp = "test",
//            Mask = "test",
//            DestinationIp = "test",
//            Vlan = "test",
//        },
//        S1C = new Route()
//        {
//            SourceIp = "test",
//            Mask = "test",
//            DestinationIp = "test",
//            Vlan = "test",
//        },
//        S1U = new Route()
//        {
//            SourceIp = "test",
//            Mask = "test",
//            DestinationIp = "test",
//            Vlan = "test",
//        },
//    },
//    TargetFile = new()
//    {
//        PathFile = "test",
//        SheetIPCLKLNK = new()
//        {
//            SheetName = "test",
//            Operation = "test",
//            Bs = "test",
//            ClientIPv4 = "test"
//        },
//        SheetVLANMAP = new()
//        {
//            SheetName = "test",
//            Operation = "test",
//            Bs = "test",
//            NextHopIP = "test",
//            Mask = "test",
//            VLANID = "test"
//        }
//    }
//};


//string fileName = "Settings.json";
//string jsonString = JsonSerializer.Serialize(settings);
//File.WriteAllText(fileName, jsonString);


//Console.WriteLine(jsonString);



string basePath = AppDomain.CurrentDomain.BaseDirectory;
string sourceFilePath = Path.Combine(basePath, "data", "Source.xlsx");
string targetFilePath = Path.Combine(basePath, "data", "Target.xlsx");
string jsonSettings = Path.Combine(basePath, "Settings.json");


#region Test



using FileStream openStream = File.OpenRead(jsonSettings);


var settings = await JsonSerializer.DeserializeAsync<Settings>(openStream);

#endregion



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