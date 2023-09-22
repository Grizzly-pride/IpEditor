﻿using IpEditor;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var sourceFilePath = @"C:\Users\Medve\Desktop\7BS_IP\Source.xlsx";
var targetFilePatch = @"C:\Users\Medve\Desktop\7BS_IP\Target.xlsx";


List<BaseStation> baseStations = await Editor.LoadSourceData(sourceFilePath);

await Editor.OpenTargetFile(targetFilePatch);


await Editor.EditOMCH(baseStations);
await Editor.EditSCTPLNK(baseStations);
await Editor.EditSCTPHOST(baseStations);
await Editor.EditUSERPLANEHOST(baseStations);
await Editor.EditIPPATH(baseStations);
await Editor.EditSRCIPRT(baseStations);
await Editor.EditDEVIP(baseStations);

Editor.CloseTargetFile();


