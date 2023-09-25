using IpEditor.Entity;
using OfficeOpenXml;
namespace IpEditor;


internal static class Editor
{
    private static ExcelPackage _package;

    public static async Task<List<BaseStation>> LoadSourceData(string filePath)
    {
        List<BaseStation> baseStations = new List<BaseStation>();

        try
        {
            var file = new FileInfo(filePath);
            if(file.Exists)
            {
                using var package = new ExcelPackage(file);
                await package.LoadAsync(file);

                var workSheet = package.Workbook.Worksheets.First();

                int row = 2;
                int column = 1;

                Logger.Info($"Loading data from a source file...");

                while (string.IsNullOrWhiteSpace(workSheet.Cells[row, column].Value?.ToString()) is false)
                {
                    string nameBS = workSheet.Cells[row, column].Value.ToString()!;

                    string sourceOAM = workSheet.Cells[row, column + 1].Value.ToString()!;
                    string nextHopOAM = workSheet.Cells[row, column + 2].Value.ToString()!;
                    string vlanOAM = workSheet.Cells[row, column + 3].Value.ToString()!;
                    string maskOAM = workSheet.Cells[row, column + 4].Value.ToString()!;
                    var oam = new Route(sourceOAM, nextHopOAM, vlanOAM, maskOAM);

                    string sourceS1C = workSheet.Cells[row, column + 5].Value.ToString()!;
                    string nextHopS1C = workSheet.Cells[row, column + 6].Value.ToString()!;
                    string vlanS1C = workSheet.Cells[row, column + 7].Value.ToString()!;
                    string maskS1C = workSheet.Cells[row, column + 8].Value.ToString()!;
                    var s1c = new Route(sourceS1C, nextHopS1C, vlanS1C, maskS1C);

                    string sourceS1U = workSheet.Cells[row, column + 5].Value.ToString()!;
                    string nextHopS1U = workSheet.Cells[row, column + 6].Value.ToString()!;
                    string vlanS1U = workSheet.Cells[row, column + 7].Value.ToString()!;
                    string maskS1U = workSheet.Cells[row, column + 8].Value.ToString()!;
                    var s1u = new Route(sourceS1U, nextHopS1U, vlanS1U, maskS1U);

                    var bs = new BaseStation(nameBS, oam, s1c, s1u);

                    baseStations.Add(bs);

                    Logger.Info($"... eNodeB: {bs.Name} has been added.");

                    row += 1;
                }
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
            bool isExistsFile = file.Exists;
            if (isExistsFile) 
            {
                _package = new ExcelPackage(file);
                await _package.LoadAsync(file);
            }
            return isExistsFile;
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

    public static void CloseTargetFile()
    {
        _package?.Dispose();
    }

    #region Edit
    public static async Task EditIPCLKLNK(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "IPCLKLNK";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, 8].Value = bs.S1C.SourceIp;
                }
                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditOMCH(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "OMCH";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, 6].Value = bs.OAM.SourceIp;
                    workSheet.Cells[row, 7].Value = bs.OAM.Mask;                  
                }
                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditSCTPLNK(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "SCTPLNK";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, 14].Value = bs.S1C.SourceIp;
                }
                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditSCTPHOST(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "SCTPHOST";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, 6].Value = bs.S1C.SourceIp;
                }
                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditUSERPLANEHOST(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "USERPLANEHOST";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, 8].Value = bs.S1U.SourceIp;
                }
                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditIPPATH(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "IPPATH";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();
                    workSheet.Cells[row, 20].Value = bs.S1U.SourceIp;
                }
                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditSRCIPRT(List<BaseStation> baseStations, string colBsName)
    {
        var nameSheet = "SRCIPRT";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[$"{colBsName}:{colBsName}"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.MOD.ToString();                   
                }
                workSheet.Cells[rows[0], 8].Value = bs.OAM.SourceIp;
                workSheet.Cells[rows[0], 10].Value = bs.OAM.DestinationIp;
                workSheet.Cells[rows[0], 14].Value = "O&M";

                workSheet.Cells[rows[1], 8].Value = bs.S1U.SourceIp;
                workSheet.Cells[rows[1], 10].Value = bs.S1U.DestinationIp;
                workSheet.Cells[rows[1], 14].Value = "S1U-MTS";

                workSheet.Cells[rows[2], 8].Value = bs.S1C.SourceIp;
                workSheet.Cells[rows[2], 10].Value = bs.S1C.DestinationIp;
                workSheet.Cells[rows[2], 14].Value = "S1C-MTS";

                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
    }

    public static async Task EditDEVIP(List<BaseStation> baseStations)
    {
        var nameSheet = "DEVIP";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        var lastUsedRow = GetLastUsedRow(workSheet);
        var firstRow = workSheet.Dimension.Start.Row;
        var firstColumn = workSheet.Dimension.Start.Column;
        var endColumn = workSheet.Dimension.End.Column;
        var workRange = workSheet.Cells[firstRow + 2, firstColumn, lastUsedRow, endColumn];

        workRange.Copy(workSheet.Cells[lastUsedRow + 1, firstColumn]);

        for ( int i = 0; i < workRange.Rows; i++ )
        {
            workRange.SetCellValue(i, 0, Operation.RMV.ToString());
        }

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[lastUsedRow + 1, 2, lastUsedRow + workRange.Rows, 2]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.ADD.ToString();
                }
                workSheet.Cells[rows[0], 10].Value = bs.OAM.SourceIp;
                workSheet.Cells[rows[0], 11].Value = bs.OAM.Mask;
                workSheet.Cells[rows[0], 12].Value = "O&M";

                workSheet.Cells[rows[1], 10].Value = bs.S1U.SourceIp;
                workSheet.Cells[rows[1], 11].Value = bs.S1U.Mask;
                workSheet.Cells[rows[1], 12].Value = "S1U-MTS";

                workSheet.Cells[rows[2], 10].Value = bs.S1C.SourceIp;
                workSheet.Cells[rows[2], 11].Value = bs.S1C.Mask;
                workSheet.Cells[rows[2], 12].Value = "S1C-MTS";

                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();       
    }

    public static async Task EditVLANMAP(List<BaseStation> baseStations)
    {
        var nameSheet = "VLANMAP";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet is null) return;

        Logger.Info($"Editing {nameSheet}...");

        var lastUsedRow = GetLastUsedRow(workSheet);
        var firstRow = workSheet.Dimension.Start.Row;
        var firstColumn = workSheet.Dimension.Start.Column;
        var endColumn = workSheet.Dimension.End.Column;
        var workRange = workSheet.Cells[firstRow + 2, firstColumn, lastUsedRow, endColumn];

        workRange.Copy(workSheet.Cells[lastUsedRow + 1, firstColumn]);

        for (int i = 0; i < workRange.Rows; i++)
        {
            workRange.SetCellValue(i, 0, Operation.RMV.ToString());
        }

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells[lastUsedRow + 1, 2, lastUsedRow + workRange.Rows, 2]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            if (rows.Any())
            {
                foreach (var row in rows)
                {
                    workSheet.Cells[row, 1].Value = Operation.ADD.ToString();
                }
                workSheet.Cells[rows[0], 4].Value = bs.OAM.DestinationIp;
                workSheet.Cells[rows[0], 5].Value = bs.OAM.Mask;
                workSheet.Cells[rows[0], 7].Value = bs.OAM.Vlan;

                workSheet.Cells[rows[1], 4].Value = bs.S1C.DestinationIp;
                workSheet.Cells[rows[1], 5].Value = bs.OAM.Mask;
                workSheet.Cells[rows[1], 7].Value = bs.OAM.Vlan;

                workSheet.Cells[rows[2], 4].Value = bs.OAM.DestinationIp;
                workSheet.Cells[rows[2], 5].Value = bs.OAM.Mask;
                workSheet.Cells[rows[2], 7].Value = bs.OAM.Vlan;

                Logger.Info($"... edited {nameSheet} for eNodeB {bs.Name} successfully.");
            }
        }
        await _package.SaveAsync();
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
