using OfficeOpenXml;

namespace IpEditor;


internal static class Editor
{
    private static ExcelPackage _package;

    public static async Task OpenTargetFile(string filePath)
    {
        try
        {
            var file = new FileInfo(filePath);
            _package = new ExcelPackage(file);
            await _package.LoadAsync(file);
        }
        catch (FileNotFoundException e)
        {
            Logger.Error($"File not found! {e.FileName}");
        }
        catch (IOException e)
        {
            Logger.Error($"File opened by another application! {e.StackTrace}");
        }
    }

    public static void CloseTargetFile()
    {
        _package?.Dispose();
    }

    #region Edit
    public static async Task EditOMCH(List<BaseStation> baseStations)
    {
        var nameSheet = "OMCH";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet == null) return; 

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells["b:b"]
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
            }
        }
        await _package.SaveAsync();

        Logger.Info($"{nameSheet} sheet has been edit successfully.");
    }

    public static async Task EditSCTPLNK(List<BaseStation> baseStations)
    {
        var nameSheet = "SCTPLNK";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet == null) return;

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells["b:b"]
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
            }
        }
        await _package.SaveAsync();

        Logger.Info($"{nameSheet} sheet has been edit successfully.");
    }

    public static async Task EditSCTPHOST(List<BaseStation> baseStations)
    {
        var nameSheet = "SCTPHOST";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet == null) return;

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells["b:b"]
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
            }
        }
        await _package.SaveAsync();

        Logger.Info($"{nameSheet} sheet has been edit successfully.");
    }

    public static async Task EditUSERPLANEHOST(List<BaseStation> baseStations)
    {
        var nameSheet = "USERPLANEHOST";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet == null) return;

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells["b:b"]
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
            }
        }
        await _package.SaveAsync();

        Logger.Info($"{nameSheet} sheet has been edit successfully.");
    }

    public static async Task EditIPPATH(List<BaseStation> baseStations)
    {
        var nameSheet = "IPPATH";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet == null) return;

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells["b:b"]
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
            }
        }
        await _package.SaveAsync();

        Logger.Info($"{nameSheet} sheet has been edit successfully.");
    }

    public static async Task EditSRCIPRT(List<BaseStation> baseStations)
    {
        var nameSheet = "SRCIPRT";
        var workSheet = GetWorkSheet(nameSheet);
        if (workSheet == null) return;

        foreach (var bs in baseStations)
        {
            var rows = workSheet!.Cells["b:b"]
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
            }
        }
        await _package.SaveAsync();

        Logger.Info($"{nameSheet} sheet has been edit successfully.");
    }
    #endregion

    public static async Task<List<BaseStation>> LoadSourceData(string filePath)
    {
        var baseStations = new List<BaseStation>();

        try
        {
            var file = new FileInfo(filePath);
            using var package = new ExcelPackage(file);
            await package.LoadAsync(file);

            var workSheet = package.Workbook.Worksheets[0];

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

                Logger.Info($"add eNodeB: {bs.Name}");

                row += 1;
            }
        }
        catch (FileNotFoundException e)
        {
            Logger.Error($"File not found! {e.FileName}");
        }
        catch (IOException e)
        {
            Logger.Error($"File opened by another application! {e.StackTrace}");
        }

        return baseStations;
    }

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
}
