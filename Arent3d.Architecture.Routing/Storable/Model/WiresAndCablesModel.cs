namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class WiresAndCablesModel
  {
    public string WireType { get ; set ; }
    public string DiameterOrNominal { get ; set ; }
    public string DOrA { get ; set ; }
    public string NumberOfHeartsOrLogarithm { get ; set ; }
    public string COrP { get ; set ; }
    public string CrossSectionalArea { get ; set ; }
    public string Name { get ; set ; }
    public string Classification { get ; set ; }
    public string FinishedOuterDiameter { get ; set ; }
    public string NumberOfConnections { get ; set ; }

    public WiresAndCablesModel( string? wireType, string? diameterOrNominal, string? dOrA, string? numberOfHeartsOrLogarithm, string? cOrP, string? crossSectionalArea, string? name, string? classification, string? finishedOuterDiameter, string? numberOfConnections )
    {
      WireType = wireType ?? string.Empty ;
      DiameterOrNominal = diameterOrNominal ?? string.Empty ;
      DOrA = dOrA ?? string.Empty ;
      COrP = cOrP ?? string.Empty ;
      NumberOfHeartsOrLogarithm = numberOfHeartsOrLogarithm ?? string.Empty ;
      CrossSectionalArea = crossSectionalArea ?? string.Empty ;
      Name = name ?? string.Empty ;
      Classification = classification ?? string.Empty ;
      FinishedOuterDiameter = finishedOuterDiameter ?? string.Empty ;
      NumberOfConnections = numberOfConnections ?? string.Empty ;
    }
  }
}