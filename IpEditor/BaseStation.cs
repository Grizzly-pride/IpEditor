namespace IpEditor;

internal sealed class BaseStation
{
    public string Name { get; set; }

    #region MME
    public string NodeOAM { get; set; }
    public string NextHopOAM { get; set; }
    public string VlanOAM { get; set; }
    public string MaskOAM { get; set; }
    #endregion

    #region S1-C
    public string NodeS1C { get; set; }
    public string NextHopS1C { get; set; }
    public string VlanS1C { get; set; }
    public string MaskS1C { get; set; }
    #endregion

    #region S1-U
    public string NodeS1U { get; set; }
    public string NextHopS1U { get; set; }
    public string VlanS1U { get; set; }
    public string MaskS1U { get; set; }
    #endregion
}
