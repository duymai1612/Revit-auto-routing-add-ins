using Autodesk.Revit.DB ;
using CsvHelper.Configuration.Attributes ;

namespace Arent3d.Architecture.Routing.AppBase
{
  /// <summary>
  /// Routing record from from-to CSV files.
  /// </summary>
  public class RouteRecord
  {
    [Index( 0 ), Name( "Route Name" )]
    public string RouteName { get ; set ; } = string.Empty ;

    [Index( 1 ), Name( "From Key" )]
    public string FromKey { get ; set ; } = string.Empty ;

    [Index( 2 ), Name( "From End Type" )]
    public string FromEndType { get ; set ; } = string.Empty ;

    [Index( 3 ), Name( "From End Param" )]
    public string FromEndParams { get ; set ; } = string.Empty ;

    [Index( 4 ), Name( "To Key" )]
    public string ToKey { get ; set ; } = string.Empty ;

    [Index( 5 ), Name( "To End Type" )]
    public string ToEndType { get ; set ; } = string.Empty ;

    [Index( 6 ), Name( "To End Param" )]
    public string ToEndParams { get ; set ; } = string.Empty ;

    [Index( 7 ), Name( "Nominal Diameter" )]
    public double? NominalDiameter { get ; set ; } = null ;

    [Index( 8 ), Name( "On Pipe Space" )]
    public bool IsRoutingOnPipeSpace { get ; set ; } = false ;

    [Index( 9 ), Name( "Preferred Curve Type Name" )]
    public string CurveTypeName { get ; set ; } = string.Empty ;

    [Index( 10 ), Name( "From Preferred Height Type" )]
    public string FromFixedHeightType { get ; set ; } = string.Empty ;

    [Index( 11 ), Name( "From Preferred Height Value" )]
    public double? FromFixedHeightValue { get ; set ; } = null ;

    [Index( 12 ), Name( "To Preferred Height Type" )]
    public string ToFixedHeightType { get ; set ; } = string.Empty ;

    [Index( 13 ), Name( "To Preferred Height Value" )]
    public double? ToFixedHeightValue { get ; set ; } = null ;

    [Index( 14 ), Name( "Preferred AvoidType" )]
    public AvoidType AvoidType { get ; set ; } = AvoidType.Whichever ;

    [Index( 15 ), Name( "MEP System Classification" )]
    public string SystemClassification { get ; set ; } = string.Empty ;

    [Index( 16 ), Name( "MEP System Type Name" )]
    public string SystemTypeName { get ; set ; } = string.Empty ;

    [Index( 17 ), Name( "Shaft Element Id" )]
    public int ShaftElementId { get ; set ; } = -1 ;
  }
}