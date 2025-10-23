using Arent3d.Revit.Csv ;
using Arent3d.Revit.Csv.Converters ;
using Autodesk.Revit.DB ;
using CsvHelper.Configuration.Attributes ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public class RackRecord
  {
    [Index( 0 ), Name( "Level" )]
    public string Level { get ; set ; } = string.Empty ;

    [Index( 1 ), Name( "OriginX" ), TypeConverter( typeof( LengthUnitConverter ) )]
    public double Origin_X { get ; set ; }

    [Index( 2 ), Name( "OriginY" ), TypeConverter( typeof( LengthUnitConverter ) )]
    public double Origin_Y { get ; set ; }

    [Index( 3 ), Name( "OriginZ" ), TypeConverter( typeof( LengthUnitConverter ) )]
    public double Origin_Z { get ; set ; }

    [Index( 4 ), Name( "RotationDegree" ), TypeConverter( typeof( LengthUnitConverter ) )]
    public double RotationDegree { get ; set ; }

    [Index( 5 ), Name( "SizeX" ), TypeConverter( typeof( LengthParameterDataConverter ) )]
    public LengthParameterData Size_X { get ; set ; } = LengthParameterData.Empty ;

    [Index( 6 ), Name( "SizeY" ), TypeConverter( typeof( LengthParameterDataConverter ) )]
    public LengthParameterData Size_Y { get ; set ; } = LengthParameterData.Empty ;

    [Index( 7 ), Name( "SizeZ" ), TypeConverter( typeof( LengthParameterDataConverter ) )]
    public LengthParameterData Size_Z { get ; set ; } = LengthParameterData.Empty ;

    [Index( 8 ), Name( "Offset" ), TypeConverter( typeof( LengthParameterDataConverter ) )]
    public LengthParameterData Offset { get ; set ; } = LengthParameterData.Empty ;

    [Index( 9 ), Name( "Elevation" ), TypeConverter( typeof( LengthParameterDataConverter ) )]
    public LengthParameterData Elevation { get ; set ; } = LengthParameterData.Empty ;

    [Ignore]
    public XYZ Origin
    {
      get => new XYZ( Origin_X, Origin_Y, Origin_Z ) ;
      set => ( Origin_X, Origin_Y, Origin_Z ) = ( value.X, value.Y, value.Z ) ;
    }

    public RackRecord()
    {
    }
  }
}