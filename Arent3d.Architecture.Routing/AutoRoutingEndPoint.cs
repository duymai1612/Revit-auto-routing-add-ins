using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Routing ;
using Arent3d.Utility ;
using MathLib ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Wrapper class for routing-core of <see cref="IAutoRoutingEndPoint"/>.
  /// </summary>
  public class AutoRoutingEndPoint : IAutoRoutingEndPoint
  {
    public IEndPoint EndPoint { get ; }

    private readonly double _minimumStraightLength ;
    private readonly double _angleToleranceRadian ;

    /// <summary>
    /// Wrap end point as an instance of <see cref="IAutoRoutingEndPoint"/>.
    /// </summary>
    /// <param name="endPoint">Base end point.</param>
    /// <param name="isFrom">True if this end point represents a from-side end point.</param>
    /// <param name="parent">An AutoRoutingEndPoint in the parent branch.</param>
    /// <param name="priority">Priority (can be duplicated between end points in an <see cref="AutoRoutingTarget"/>).</param>
    /// <param name="edgeDiameter">Edge diameter.</param>
    /// <param name="isDirect">Whether use direct routing or not.</param>
    /// <param name="routeCondition"></param>
    public AutoRoutingEndPoint( IEndPoint endPoint, bool isFrom, IAutoRoutingEndPoint? parent, int priority, double edgeDiameter, bool isDirect, MEPSystemRouteCondition routeCondition )
    {
      EndPoint = endPoint ;
      IsStart = isFrom ;
      Priority = priority ;
      Parent = parent ;
      Depth = priority ;
      IsDirect = isDirect ;

      var pipeSpec = routeCondition.Spec ;
      _minimumStraightLength = pipeSpec.GetReducerLength( (endPoint.GetDiameter() ?? -1).DiameterValueToPipeDiameter(), edgeDiameter.DiameterValueToPipeDiameter() ) + EndPoint.GetMinimumStraightLength( edgeDiameter, IsStart ) ;
      _angleToleranceRadian = pipeSpec.AngleTolerance ;

      PipeCondition = routeCondition ;
    }

    public bool IsParentOf( IAutoRoutingEndPoint ep )
    {
      var targetParent = ep.Parent is IPseudoEndPoint pseudoParent ? pseudoParent.Source : ep.Parent ;
      return targetParent == this ;
    }

    public IAutoRoutingEndPoint? Parent { get ; }

    public Vector3d Position => EndPoint.RoutingStartPosition.To3dRaw() + _minimumStraightLength * Direction.ForEndPointType( IsStart ) ;

    public Vector3d Direction => Sanitize( EndPoint.GetRoutingDirection( IsStart ).To3dRaw() ) ;
    public Vector3d? PositionConstraint => null ;

    private static readonly Vector3d[] SanitizationDirections =
    {
      new Vector3d( +1, 0, 0 ), new Vector3d( -1, 0, 0 ),
      new Vector3d( 0, +1, 0 ), new Vector3d( 0, -1, 0 ),
      new Vector3d( 0, 0, +1 ), new Vector3d( 0, 0, -1 ),
    } ;
    private Vector3d Sanitize( Vector3d vec )
    {
      foreach ( var dir in SanitizationDirections ) {
        double dot = Vector3d.Dot( vec, dir ), cross = Vector3d.Cross( vec, dir ).magnitude ;
        if ( Math.Atan2( cross, dot ) < _angleToleranceRadian ) return dir ;
      }

      return vec ;
    }

    /// <summary>
    /// Returns a routing condition object determined from the related connector.
    /// </summary>
    public IRouteCondition PipeCondition { get ; }

    /// <summary>
    /// Returns whether this end point is from-side end point.
    /// </summary>
    public bool IsStart { get ; }

    /// <summary>
    /// Returns the priority. <see cref="Priority"/> is similar to <see cref="Depth"/>, but can be duplicated between end points in an <see cref="AutoRoutingTarget"/>.
    /// </summary>
    public int Priority { get ; }

    /// <summary>
    /// Returns the priority. <see cref="Depth"/> is similar to <see cref="Priority"/>, but cannot be duplicated between end points in an <see cref="AutoRoutingTarget"/>.
    /// </summary>
    public int Depth { get ; }

    public bool AllowHorizontalBranches => true ;
    public bool IsDirect { get ; }
    public bool AllowThroughBatteryLimit => false ;

    /// <summary>
    /// Returns this end point's floating type. Now it always returns <see cref="RoutingPointType.OtherNozzle"/> (i.e. non-floated).
    /// </summary>
    public RoutingPointType PointType => RoutingPointType.OtherNozzle ;

    /// <summary>
    /// Not used now. Always returns null.
    /// </summary>
    public ILayerStack? LinkedRack => null ;
  }
}