using System ;
using System.Diagnostics ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing.EndPoints
{
  [DebuggerDisplay("{Type}:{Key}")]
  public class DummyBreakEndPoint : IEndPoint
  {
    public DummyBreakEndPoint( Vector3d breakPosition, Vector3d routingDirection, int id )
    {
      RoutingStartPosition = breakPosition.ToXYZRaw() ;
      RoutingDirection = routingDirection.ToXYZRaw() ;
      Id = id ;
    }

    public const string Type = "Break point" ;

    public int Id { get ; }
    
    public string TypeName => Type ;
    public string DisplayTypeName => Type ;
    public EndPointKey Key => new EndPointKey( Type, ParameterString ) ;
    public bool IsReplaceable => false ;
    public bool IsOneSided => false ;
    public string ParameterString => Id.ToString() ;

    public XYZ RoutingStartPosition { get ; }
    public XYZ RoutingDirection { get ; }

    public XYZ GetRoutingDirection( bool isFrom ) => RoutingDirection ;

    public double? GetDiameter() => null ;

    public double GetMinimumStraightLength( double edgeDiameter, bool isFrom ) => 0 ;

    public ElementId GetLevelId( Document document ) => ElementId.InvalidElementId ;

    Connector? IEndPoint.GetReferenceConnector() => throw new NotSupportedException() ;
    bool IEndPoint.HasValidElement( bool isFrom ) => throw new NotSupportedException() ;
    Route? IEndPoint.ParentRoute() => throw new NotSupportedException() ;
    SubRoute? IEndPoint.ParentSubRoute() => throw new NotSupportedException() ;
    bool IEndPoint.GenerateInstance( string routeName ) => throw new NotSupportedException() ;
    bool IEndPoint.EraseInstance() => throw new NotSupportedException() ;
    void IEndPoint.Accept( IEndPointVisitor visitor ) => throw new NotSupportedException() ;
    T IEndPoint.Accept<T>( IEndPointVisitor<T> visitor ) => throw new NotSupportedException() ;
  }
}