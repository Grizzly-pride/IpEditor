using OfficeOpenXml;
namespace IpEditor;


internal static class Editor
{
    private static ExcelPackage _package;

    public static async Task OpenTargetFile(string filePath)
    {
        var file = new FileInfo(filePath);
        _package = new ExcelPackage(file);
        await _package.LoadAsync(file);
    }

    public static void CloseTargetFile()
    {
        _package?.Dispose();
    }

    public static async Task EditOMCH(List<BaseStation> baseStations)
    {
        string sheetName = "OMCH";
        var workSheet = _package.Workbook.Worksheets.First(a => a.Name.Equals(sheetName))
            ?? throw new NullReferenceException($"{sheetName} not found !!!");

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
    }

    public static async Task<List<BaseStation>> LoadSourceData(string filePath)
    {
        var file = new FileInfo(filePath);
        using var package = new ExcelPackage(file);
        await package.LoadAsync(file);

        var workSheet = package.Workbook.Worksheets[0];

        int row = 2;
        int column = 1;

        var baseStations = new List<BaseStation>();

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

            Console.WriteLine($"Get eNodeB: {bs.Name}");

            row += 1;
        }

        return baseStations;
    }
}
