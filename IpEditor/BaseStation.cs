namespace IpEditor;

internal sealed class BaseStation
{
    public string Name { get; }
    public Route OAM { get; }
    public Route S1C { get; }
    public Route S1U { get; }

    public BaseStation(string name, Route oam, Route s1c, Route s1u)
    {
        Name = name;
        OAM = oam;
        S1C = s1c;
        S1U = s1u;           
    }
}
