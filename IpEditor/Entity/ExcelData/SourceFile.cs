namespace IpEditor.Entity.ExcelData;

internal sealed class SourceFile
{
    public string PathFile { get; init; }
    public string BSName { get; init; }
    public Route OAM { get; init; }
    public Route S1C { get; init; }
    public Route S1U { get; init;}
}
