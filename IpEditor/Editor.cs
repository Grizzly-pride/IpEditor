using IpEditor.Entity;
using IpEditor.Entity.ExcelData;
using IpEditor.Entity.ExcelData.Sheets;
using OfficeOpenXml;
namespace IpEditor;


internal static class Editor
{
    private static ExcelPackage _package;

    public static async Task<List<BaseStation>> LoadSourceData(SourceFile sourceFile)
    {
        string sourceFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data", "Source.xlsx");
        List<BaseStation> baseStations = new List<BaseStation>();

        try
        {
            var file = new FileInfo(sourceFile.PathFile ?? sourceFilePath);
           
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);

            var workSheet = package.Workbook.Worksheets.First();

            #region Columns
            int startRow = 2;

            int colBsName = sourceFile.BSName.ExcelColNameToInt();

            int colBsOAM = sourceFile.OAM.SourceIp.ExcelColNameToInt();
            int colGatewayOAM = sourceFile.OAM.DestinationIp.ExcelColNameToInt();
            int colVlanOAM = sourceFile.OAM.Vlan.ExcelColNameToInt();
            int colMaskOAM = sourceFile.OAM.Mask.ExcelColNameToInt();
            int colLabelOAM = sourceFile.OAM.Label.ExcelColNameToInt();

            int colBsS1U = sourceFile.S1U.SourceIp.ExcelColNameToInt();
            int colGatewayS1U = sourceFile.S1U.DestinationIp.ExcelColNameToInt();
            int colVlanS1U = sourceFile.S1U.Vlan.ExcelColNameToInt();
            int colMaskS1U = sourceFile.S1U.Mask.ExcelColNameToInt();
            int colLabelS1U = sourceFile.S1U.Label.ExcelColNameToInt();

            int colBsS1C = sourceFile.S1C.SourceIp.ExcelColNameToInt();
            int colGatewayS1C = sourceFile.S1C.DestinationIp.ExcelColNameToInt();
            int colVlanS1C = sourceFile.S1C.Vlan.ExcelColNameToInt();
            int colMaskS1C = sourceFile.S1C.Mask.ExcelColNameToInt();
            int colLabelS1C = sourceFile.S1C.Label.ExcelColNameToInt();
            #endregion

            Logger.Info($"[ Loading data from a source file ]");

            while (string.IsNullOrWhiteSpace(workSheet.Cells[startRow, colBsName].Value?.ToString()) is false)
            {
                string nameBS = workSheet.Cells[startRow, colBsName].Value.ToString()!;

                string sourceOAM = workSheet.Cells[startRow, colBsOAM].Value.ToString()!;
                string nextHopOAM = workSheet.Cells[startRow, colGatewayOAM].Value.ToString()!;
                string vlanOAM = workSheet.Cells[startRow, colVlanOAM].Value.ToString()!;
                string maskOAM = workSheet.Cells[startRow, colMaskOAM].Value.ToString()!;
                string labelOAM = workSheet.Cells[startRow, colLabelOAM].Value.ToString()!;
                var oam = new Route(sourceOAM, nextHopOAM, vlanOAM, maskOAM, labelOAM);

                string sourceS1U = workSheet.Cells[startRow, colBsS1U].Value.ToString()!;
                string nextHopS1U = workSheet.Cells[startRow, colGatewayS1U].Value.ToString()!;
                string vlanS1U = workSheet.Cells[startRow, colVlanS1U].Value.ToString()!;
                string maskS1U = workSheet.Cells[startRow, colMaskS1U].Value.ToString()!;
                string labelS1U = workSheet.Cells[startRow, colLabelS1U].Value.ToString()!;
                var s1u = new Route(sourceS1U, nextHopS1U, vlanS1U, maskS1U, labelS1U);

                string sourceS1C = workSheet.Cells[startRow, colBsS1C].Value.ToString()!;
                string nextHopS1C = workSheet.Cells[startRow, colGatewayS1C].Value.ToString()!;
                string vlanS1C = workSheet.Cells[startRow, colVlanS1C].Value.ToString()!;
                string maskS1C = workSheet.Cells[startRow, colMaskS1C].Value.ToString()!;
                string labelS1C = workSheet.Cells[startRow, colLabelS1C].Value.ToString()!;
                var s1c = new Route(sourceS1C, nextHopS1C, vlanS1C, maskS1C, labelS1C);

                var bs = new BaseStation(nameBS, oam, s1c, s1u);

                baseStations.Add(bs);

                Logger.Info($"... eNodeB: {bs.Name} has been added.");

                startRow += 1;
            }
            
        }
        catch (Exception e)
        {
            if(e is FileNotFoundException fileNotFound)
            {
                Logger.Error($"File not found! {fileNotFound.FileName}");
            }
            else if(e is IOException io)
            {
                Logger.Error($"File opened by another application! {io.Source}");
            }

            Logger.Error($"{e.StackTrace}");
        }

        return baseStations;
    }

    public static async Task<bool> OpenTargetFile(string filePath)
    {
        try
        {
            var file = new FileInfo(filePath);
            _package = new ExcelPackage(file);
            await _package.LoadAsync(file);

            return file.Exists;
        }
        catch(Exception e)
        {
            if (e is FileNotFoundException fileNotFound)
            {
                Logger.Error($"File not found! {fileNotFound.FileName}");
            }
            else if (e is IOException io)
            {
                Logger.Error($"File opened by another application! {io.Source}");
            }

            Logger.Error($"{e.StackTrace}");

            return false;
        }
    }

    public static async Task CloseTargetFile()
    {
        try
        {
            await _package.SaveAsync();
        }
        catch (InvalidOperationException e) 
        {
            Logger.Error(e.Message);
        }
        finally
        {
            _package?.Dispose();
        }
    }

    #region Edit
    public static void EditIPCLKLNK(List<BaseStation> baseStations, SheetIPCLKLNK sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;
        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIPv4 = sheet.ClientIPv4!.ExcelColNameToInt();


        Logger.Info($"[ Editing {sheet.SheetName} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, colIPv4].Value = bs.S1C.SourceIp;
                }
                Logger.Info($"... edited {sheet.SheetName} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditOMCH(List<BaseStation> baseStations, SheetOMCH sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIp = sheet.LocalIP.ExcelColNameToInt();
        int colMask = sheet.LocalMask.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, colIp].Value = bs.OAM.SourceIp;
                    workSheet.Cells[row, colMask].Value = bs.OAM.Mask;                  
                }
                Logger.Info($"... edited {sheet.SheetName} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditSCTPLNK(List<BaseStation> baseStations, SheetSCTPLNK sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIp = sheet.FirstLocalIPAddress!.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, colIp].Value = bs.S1C.SourceIp;
                }
                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditSCTPHOST(List<BaseStation> baseStations, SheetSCTPHOST sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIp = sheet.FirstLocalIPAddress.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, colIp].Value = bs.S1C.SourceIp;
                }
                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditUSERPLANEHOST(List<BaseStation> baseStations, SheetUSERPLANEHOST sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIp = sheet.LocalIPAddress!.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, colIp].Value = bs.S1U.SourceIp;
                }
                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditIPPATH(List<BaseStation> baseStations, SheetIPPATH sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIp = sheet.LocalIPAddress!.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, colIp].Value = bs.S1U.SourceIp;
                }
                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditSRCIPRT(List<BaseStation> baseStations, SheetSRCIPRT sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colSourceIp = sheet.SourceIPAddress!.ExcelColNameToInt();
        int colGatewayIp = sheet.NextHopIP!.ExcelColNameToInt();
        int colLabel = sheet.UserLabel.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{sheet.Bs}:{sheet.Bs}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.MOD.ToString();                   
                }
                workSheet.Cells[rows[0], colSourceIp].Value = bs.OAM.SourceIp;
                workSheet.Cells[rows[0], colGatewayIp].Value = bs.OAM.DestinationIp;
                workSheet.Cells[rows[0], colLabel].Value = bs.OAM.Label;

                workSheet.Cells[rows[1], colSourceIp].Value = bs.S1U.SourceIp;
                workSheet.Cells[rows[1], colGatewayIp].Value = bs.S1U.DestinationIp;
                workSheet.Cells[rows[1], colLabel].Value = bs.S1U.Label;

                workSheet.Cells[rows[2], colSourceIp].Value = bs.S1C.SourceIp;
                workSheet.Cells[rows[2], colGatewayIp].Value = bs.S1C.DestinationIp;
                workSheet.Cells[rows[2], colLabel].Value = bs.S1C.Label;

                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }
    }

    public static void EditDEVIP(List<BaseStation> baseStations, SheetDEVIP sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colBs = sheet.Bs!.ExcelColNameToInt(); 
        int colIp = sheet.IPAddress.ExcelColNameToInt();
        int colM = sheet.Mask.ExcelColNameToInt();
        int colU = sheet.UserLabel.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        var lastUsedRow = GetLastUsedRow(workSheet);
        var firstRow = workSheet.Dimension.Start.Row;
        var firstColumn = workSheet.Dimension.Start.Column;
        var endColumn = workSheet.Dimension.End.Column;
        var workRange = workSheet.Cells[firstRow + 2, firstColumn, lastUsedRow, endColumn];

        workRange.Copy(workSheet.Cells[lastUsedRow + 1, firstColumn]);

        var colOperIndex = colOper - 1;
        for ( int i = 0; i < workRange.Rows; i++ )
        {
            workRange.SetCellValue(i, colOperIndex, Operation.RMV.ToString());
        }

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[lastUsedRow + 1, colBs, lastUsedRow + workRange.Rows, colBs]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.ADD.ToString();
                }
                workSheet.Cells[rows[0], colIp].Value = bs.OAM.SourceIp;
                workSheet.Cells[rows[0], colM].Value = bs.OAM.Mask;
                workSheet.Cells[rows[0], colU].Value = bs.OAM.Label;

                workSheet.Cells[rows[1], colIp].Value = bs.S1U.SourceIp;
                workSheet.Cells[rows[1], colM].Value = bs.S1U.Mask;
                workSheet.Cells[rows[1], colU].Value = bs.S1U.Label;

                workSheet.Cells[rows[2], colIp].Value = bs.S1C.SourceIp;
                workSheet.Cells[rows[2], colM].Value = bs.S1C.Mask;
                workSheet.Cells[rows[2], colU].Value = bs.S1C.Label;

                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }      
    }

    public static void EditVLANMAP(List<BaseStation> baseStations, SheetVLANMAP sheet)
    {
        var workSheet = GetWorkSheet(sheet.SheetName!);
        if (workSheet is null) return;

        int colBs = sheet.Bs!.ExcelColNameToInt();
        int colOper = sheet.Operation!.ExcelColNameToInt();
        int colIp = sheet.NextHopIP!.ExcelColNameToInt();
        int colM = sheet.Mask!.ExcelColNameToInt();
        int colV = sheet.VLANID!.ExcelColNameToInt();

        Logger.Info($"[ Editing {sheet.SheetName!} ]");

        var lastUsedRow = GetLastUsedRow(workSheet);
        var firstRow = workSheet.Dimension.Start.Row;
        var firstColumn = workSheet.Dimension.Start.Column;
        var endColumn = workSheet.Dimension.End.Column;
        var workRange = workSheet.Cells[firstRow + 2, firstColumn, lastUsedRow, endColumn];

        workRange.Copy(workSheet.Cells[lastUsedRow + 1, firstColumn]);

        var colOperIndex = colOper - 1;
        for (int i = 0; i < workRange.Rows; i++)
        {
            workRange.SetCellValue(i, colOperIndex, Operation.RMV.ToString());
        }

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[lastUsedRow + 1, colBs, lastUsedRow + workRange.Rows, colBs]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, colOper].Value = Operation.ADD.ToString();
                }
                workSheet.Cells[rows[0], colIp].Value = bs.OAM.DestinationIp;
                workSheet.Cells[rows[0], colM].Value = bs.OAM.Mask;
                workSheet.Cells[rows[0], colV].Value = bs.OAM.Vlan;

                workSheet.Cells[rows[1], colIp].Value = bs.S1C.DestinationIp;
                workSheet.Cells[rows[1], colM].Value = bs.OAM.Mask;
                workSheet.Cells[rows[1], colV].Value = bs.OAM.Vlan;

                workSheet.Cells[rows[2], colIp].Value = bs.OAM.DestinationIp;
                workSheet.Cells[rows[2], colM].Value = bs.OAM.Mask;
                workSheet.Cells[rows[2], colV].Value = bs.OAM.Vlan;

                Logger.Info($"... edited {sheet.SheetName!} for eNodeB {bs.Name} successfully.");
            }
        }
    }
    #endregion

    private static ExcelWorksheet? GetWorkSheet(string sheetName)
    {
        ExcelWorksheet? workSheet = null;
        try
        {
            workSheet = _package.Workbook.Worksheets
                .First(a => a.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            return workSheet;

        }
        catch (InvalidOperationException) 
        {
            Logger.Warning($"Sheet {sheetName} not found!");
        }
        return workSheet;     
    }

    private static int GetLastUsedRow(ExcelWorksheet sheet)
    {
        if (sheet.Dimension == null) return default;

        var row = sheet.Dimension.End.Row;
        while (row > 0)
        {
            var range = sheet.Cells[row, sheet.Dimension.Start.Column, row, sheet.Dimension.End.Column];
            if (range.Any(c => !string.IsNullOrEmpty(c.Text)))
            {
                break;
            }
            row--;
        }
        return row;
    }
}
