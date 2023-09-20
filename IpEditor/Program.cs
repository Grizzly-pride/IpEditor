using IpEditor;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

Editor.PathSourceExcelFile = @"C:\Users\Medve\Desktop\7BS_IP\Source.xlsx";
Editor.PathTargetExcelFile = @"C:\Users\Medve\Desktop\7BS_IP\Target.xlsx";

List<BaseStation> baseStations = await Editor.GetSourceData();

await Editor.EditOMCH(baseStations);
