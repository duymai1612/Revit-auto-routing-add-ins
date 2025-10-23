namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class ConduitsModel
  {
    public string PipingType { get ; set ; }
    public string Size { get ; set ; }
    public string InnerCrossSectionalArea { get ; set ; }
    public string Name { get ; set ; }
    public string Classification { get ; set ; }

    public ConduitsModel( string? pipingType, string? size, string? innerCrossSectionalArea, string? name, string? classification )
    {
      PipingType = pipingType ?? string.Empty ;
      Size = size ?? string.Empty ;
      InnerCrossSectionalArea = innerCrossSectionalArea ?? string.Empty ;
      Name = name ?? string.Empty ;
      Classification = classification ?? string.Empty ;
    }
  }
}