using System ;
using Arent3d.Architecture.Routing.Utils ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class HeightSettingModel : IEquatable<HeightSettingModel>
  {
    private const double DEFAULT_HEIGHT_OF_LEVEL = 3000 ;
    private const double DEFAULT_HEIGHT_OF_CONNECTORS = 2000 ;
    private const double DEFAULT_UNDERFLOOR = -500 ;
    private const string DEFAULT_LEVEL_NAME = "(No level name)" ;

    public int LevelId { get ; set ; }
    public string LevelName { get ; set ; }
    public double Elevation { get ; set ; }
    public double Underfloor { get ; set ; }
    public double HeightOfLevel { get ; set ; }
    public double HeightOfConnectors { get ; set ; }

    public HeightSettingModel( Level levels )
    {
      if ( levels == null )
        throw new ArgumentNullException( nameof( levels ) ) ;

      LevelId = levels.Id.IntegerValue ;
      LevelName = StringUtils.DefaultIfBlank( levels.Name, DEFAULT_LEVEL_NAME ) ;
      Elevation = Math.Round(levels.Elevation.RevitUnitsToMillimeters()) ;
      Underfloor = DEFAULT_UNDERFLOOR ;
      HeightOfLevel = DEFAULT_HEIGHT_OF_LEVEL ;
      HeightOfConnectors = DEFAULT_HEIGHT_OF_CONNECTORS ;
    }

    public HeightSettingModel( Level levels, double elevation, double underfloor, double heightOfLevel, double heightOfConnectors )
    {
      if ( levels == null )
        throw new ArgumentNullException( nameof( levels ) ) ;

      LevelId = levels.Id.IntegerValue ;
      LevelName = StringUtils.DefaultIfBlank( levels.Name, DEFAULT_LEVEL_NAME ) ;
      Elevation = Math.Round( elevation ) ;
      Underfloor = Math.Round( underfloor ) ;
      HeightOfLevel = Math.Round( heightOfLevel ) ;
      HeightOfConnectors = Math.Round( heightOfConnectors ) ;
    }

    public HeightSettingModel( int? levelId, string? levelName, double? elevation, double? underfloor, double? heightOfLevel, double? heightOfConnectors )
    {
      LevelId = levelId ?? throw new ArgumentNullException( nameof( levelId ) ) ;
      LevelName = StringUtils.DefaultIfBlank( levelName, DEFAULT_LEVEL_NAME ) ;
      Elevation = Math.Round( elevation ?? 0 ) ;
      Underfloor = Math.Round( underfloor ?? DEFAULT_UNDERFLOOR ) ;
      HeightOfLevel = Math.Round( heightOfLevel ?? DEFAULT_HEIGHT_OF_LEVEL ) ;
      HeightOfConnectors = Math.Round( heightOfConnectors ?? DEFAULT_HEIGHT_OF_CONNECTORS ) ;
    }

    public bool Equals( HeightSettingModel other )
    {
      return other != null &&
             LevelId == other.LevelId &&
             LevelName == other.LevelName &&
             Elevation == other.Elevation &&
             Underfloor == other.Underfloor &&
             HeightOfLevel == other.HeightOfLevel &&
             HeightOfConnectors == other.HeightOfConnectors ;
    }
  }
}