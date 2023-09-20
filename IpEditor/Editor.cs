using OfficeOpenXml;
namespace IpEditor;

internal static class Editor
{
    public static string PathSourceExcelFile { get; set; }
    public static string PathTargetExcelFile { get; set; }


    public static async Task EditOMCH(List<BaseStation> baseStations)
    {

    }

    public static async Task<List<BaseStation>> GetSourceData()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var file = new FileInfo(PathSourceExcelFile);
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

            row += 1;
        }

        return baseStations;
    }
}
