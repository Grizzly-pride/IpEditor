using IpEditor;
using OfficeOpenXml;



ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var sourceFilePath = @"C:\Users\alexandr.medved\Desktop\Test\Source.xlsx";
var targetFilePatch = @"C:\Users\alexandr.medved\Desktop\Test\Target.xlsx";


List<BaseStation> baseStations = await Editor.LoadSourceData(new FileInfo(sourceFilePath));

await Editor.OpenTargetFile(new FileInfo(targetFilePatch));
await Editor.EditOMCH(baseStations);
