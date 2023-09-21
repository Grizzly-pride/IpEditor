namespace IpEditor;

internal sealed class Route
{
    public string SourceIp { get; }
    public string DestinationIp { get; }
    public string Vlan { get; }
    public string Mask { get; }

    public Route(string sourceIp, string destinationIp, string vlan, string mask)
    {
        SourceIp = sourceIp;
        DestinationIp = destinationIp;
        Vlan = vlan;
        Mask = mask;           
    }
}
