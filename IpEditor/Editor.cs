using OfficeOpenXml;
namespace IpEditor;

internal static class Editor
{
    private static ExcelPackage _package;

    public static async Task OpenTargetFile(FileInfo file)
    {
        using var _package = new ExcelPackage(file);
        await _package.LoadAsync(file);
    }


    public static async Task EditOMCH(List<BaseStation> baseStations)
    {
        //var file = new FileInfo(PathTargetExcelFile);
        //using var package = new ExcelPackage(file);
        //await package.LoadAsync(file);

        var workSheet = _package.Workbook.Worksheets.First(a => a.Name.Equals("OMCH"));

        foreach (var bs in baseStations)
        {
            var rows = workSheet.Cells["b:b"]
                .Where(cel => cel.Text.StartsWith(bs.Name, StringComparison.OrdinalIgnoreCase))
                .Select(i => i.End.Row)
                .ToList();

            foreach(var row in rows)
            {
                
            }
        }

        //var rows = workSheet.Cells["a:a"].GetCellValue<string>(1, 0);

        //int rowCount = workSheet.Dimension.End.Row;

        //int count = 0;

        //for (int row = 3; row <= rowCount - 1; row++)
        //{
        //    var rowVal = workSheet.Cells[row, 2].Value.ToString();

        //    if (rowVal.Equals("705621_Osovo"))
        //    {
        //         count++;
        //    }
        //};

        //var query = workSheet.Cells["a:a"]
        //    .Where(cel => cel.Value == ("Lerka"))
        //    .Select(cel => cel.Address).ToList();

        //var query = from cell in workSheet.Cells["b:b"]
        //            where cell.Value.ToString() == "Lerka"
        //            select cell.Rows;
        //var list = query.ToList();
        //foreach(var cell in query)
        //{
        //    Console.WriteLine($"{cell}");
        //}

        //int rowNum = searchCell.First();

        //searchCell.ToList().ForEach(i => Console.WriteLine(i));
    }

    public static async Task<List<BaseStation>> LoadSourceData(FileInfo file)
    {
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
