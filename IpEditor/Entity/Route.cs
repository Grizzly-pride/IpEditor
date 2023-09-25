﻿namespace IpEditor.Entity;

internal sealed class Route
{
    public string SourceIp { get; init; }
    public string DestinationIp { get; init; }
    public string Vlan { get; init; }
    public string Mask { get; init; }

    public Route() { }

    public Route(string sourceIp, string destinationIp, string vlan, string mask)
    {
        SourceIp = sourceIp;
        DestinationIp = destinationIp;
        Vlan = vlan;
        Mask = mask;
    }
}
