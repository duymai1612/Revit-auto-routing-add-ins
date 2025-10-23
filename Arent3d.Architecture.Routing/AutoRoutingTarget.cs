using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.Extensions ;
using Arent3d.Routing ;
using Arent3d.Routing.Conditions ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing
{
  public class AutoRoutingTarget : IAutoRoutingTarget
  {
    /// <summary>
    /// All routes an <see cref="AutoRoutingTarget"/> is related to.
    /// </summary>
    public IReadOnlyCollection<Route> Routes { get ; }

    /// <summary>
    /// Routing end points which fluid flows from.
    /// </summary>
    private readonly IReadOnlyCollection<AutoRoutingEndPoint> _fromEndPoints ;

    /// <summary>
    /// Routing end points which fluid flows to.
    /// </summary>
    private readonly IReadOnlyCollection<AutoRoutingEndPoint> _toEndPoints ;

    private readonly IReadOnlyDictionary<AutoRoutingEndPoint, SubRoute> _ep2SubRoute ;

    /// <summary>
    /// Returns all routing end points.
    /// </summary>
    public IEnumerable<AutoRoutingEndPoint> EndPoints => _fromEndPoints.Concat( _toEndPoints ) ;

    public Domain Domain { get ; }

    public AutoRoutingTarget( Document document, IReadOnlyCollection<SubRoute> subRoutes, IReadOnlyDictionary<Route, int> priorities, IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> routeConditionDictionary )
    {
      if ( 0 == subRoutes.Count ) throw new ArgumentException() ;

      Routes = subRoutes.Select( subRoute => subRoute.Route ).Distinct().EnumerateAll() ;
      Domain = Routes.Select( route => route.Domain ).First() ;

      var depths = GetDepths( subRoutes ) ;
      var ep2SubRoute = new Dictionary<AutoRoutingEndPoint, SubRoute>() ;
      var fromEndPoints = new List<AutoRoutingEndPoint>() ;
      var toEndPoints = new List<AutoRoutingEndPoint>() ;
      var subRoute2ToEndPoint = new Dictionary<SubRoute, AutoRoutingEndPoint>() ;

      foreach ( var subRoute in subRoutes.OrderBy( s => depths[ s ] ) ) {
        var depth = depths[ subRoute ] ;
        var diameter = subRoute.GetDiameter() ;
        var isDirect = ( false == subRoute.IsRoutingOnPipeSpace ) ;
        var routeCondition = routeConditionDictionary[ new SubRouteInfo( subRoute ) ] ;

        var routeEndPoints = subRoute.AllEndPoints.OfType<RouteEndPoint>().ToList() ;
        if ( 2 <= routeEndPoints.Count ) {
          throw new InvalidOperationException( "RouteEndPoint must be one or less in an AutoRoutingTarget." ) ;
        }

        AutoRoutingEndPoint? parent ;
        if ( 1 == routeEndPoints.Count ) {
          var parentSubRoute = routeEndPoints[ 0 ].ParentSubRoute() ;
          if ( null == parentSubRoute || false == subRoute2ToEndPoint.TryGetValue( parentSubRoute, out var ep ) ) {
            throw new InvalidOperationException( "RouteEndPoint's parent is not contained in its owner AutoRoutingTarget." ) ;
          }
          parent = ep ;
        }
        else {
          parent = null ;
        }

        foreach ( var (endPoints, isFrom) in new[] { ( subRoute.FromEndPoints, true ), ( subRoute.ToEndPoints, false ) } ) {
          foreach ( var endPoint in endPoints.Where( ep => ep is not RouteEndPoint ) ) {
            var ep = new AutoRoutingEndPoint( endPoint, isFrom, parent, depth, diameter, isDirect, routeCondition ) ;
            ep2SubRoute.Add( ep, subRoute ) ;
            if ( isFrom ) {
              fromEndPoints.Add( ep ) ;
            }
            else {
              toEndPoints.Add( ep ) ;
            }

            if ( isFrom && null == parent ) continue ;
            if ( subRoute2ToEndPoint.ContainsKey( subRoute ) ) continue ;

            subRoute2ToEndPoint.Add( subRoute, ep ) ;
          }
        }
      }

      _fromEndPoints = fromEndPoints ;
      _toEndPoints = toEndPoints ;
      _ep2SubRoute = ep2SubRoute ;

      var firstSubRoute = subRoutes.First() ;
      LineId = $"{firstSubRoute.Route.RouteName}@{firstSubRoute.SubRouteIndex}" ;

      var trueFixedBopHeight = firstSubRoute.GetTrueFixedBopHeight( FixedHeightUsage.Default ) ;
      Condition = new AutoRoutingCondition( document, firstSubRoute, priorities[ firstSubRoute.Route ], trueFixedBopHeight ) ;
    }

    public AutoRoutingTarget( Document document, SubRoute subRoute, int priority, AutoRoutingEndPoint fromEndPoint, AutoRoutingEndPoint toEndPoint, double? forcedFixedHeight )
    {
      Routes = new[] { subRoute.Route } ;
      Domain = subRoute.Route.Domain ;

      _fromEndPoints = new[] { fromEndPoint } ;
      _toEndPoints = new[] { toEndPoint } ;
      _ep2SubRoute = new Dictionary<AutoRoutingEndPoint, SubRoute> { { fromEndPoint, subRoute }, { toEndPoint, subRoute } } ;

      LineId = $"{subRoute.Route.RouteName}@{subRoute.SubRouteIndex}" ;

      Condition = new AutoRoutingCondition( document, subRoute, priority, forcedFixedHeight ) ;
    }

    private static IReadOnlyDictionary<SubRoute, int> GetDepths( IReadOnlyCollection<SubRoute> subRoutes )
    {
      var parentInfo = CollectSubRouteParents( subRoutes ) ;

      var result = new Dictionary<SubRoute, int>() ;
      var newDepthList = new List<SubRoute>() ;
      var newDepth = 0 ;
      while ( 0 < parentInfo.Count ) {
        newDepthList.Clear() ;

        foreach ( var (subRoute, parents) in parentInfo ) {
          if ( 0 == parents.Count ) {
            newDepthList.Add( subRoute ) ;
          }
        }

        foreach ( var subRoute in newDepthList ) {
          result.Add( subRoute, newDepth ) ;

          parentInfo.Remove( subRoute ) ;
        }

        foreach ( var (_, parents) in parentInfo ) {
          parents.RemoveAll( newDepthList.Contains ) ;
        }

        ++newDepth ;
      }

      return result ;
    }

    private static Dictionary<SubRoute, List<SubRoute>> CollectSubRouteParents( IReadOnlyCollection<SubRoute> subRoutes )
    {
      var subRouteAndParents = subRoutes.ToDictionary( r => r, r => new List<SubRoute>() ) ;

      foreach ( var subRoute in subRoutes ) {
        foreach ( var parent in subRoute.AllEndPoints.Select( ep => ep.ParentSubRoute() ).NonNull() ) {
          if ( false == subRouteAndParents.TryGetValue( subRoute, out var list ) ) continue ; // not contained.

          list.Add( parent ) ;
        }
      }

      return subRouteAndParents ;
    }

    public IAutoRoutingSpatialConstraints? CreateConstraints()
    {
      if ( ( 0 < _fromEndPoints.Count ) && ( 0 < _toEndPoints.Count ) ) {
        return new AutoRoutingSpatialConstraints( _fromEndPoints, _toEndPoints ) ;
      }

      return null ;
    }

    public string LineId { get ; }

    public ICommonRoutingCondition Condition { get ; }

    public int RouteCount => _fromEndPoints.Count + _toEndPoints.Count - 1 ;

    public Action<IEnumerable<(IAutoRoutingEndPoint, Vector3d)>> PositionInitialized => SyncTermPositions ;

    private static void SyncTermPositions( IEnumerable<(IAutoRoutingEndPoint, Vector3d)> positions )
    {
      foreach ( var (autoRoutingEndPoint, position) in positions ) {
        if ( autoRoutingEndPoint is not AutoRoutingEndPoint endPoint ) throw new Exception() ;

        // do nothing now.
      }
    }

    public SubRoute? GetSubRoute( AutoRoutingEndPoint ep )
    {
      return _ep2SubRoute.TryGetValue( ep, out var subRoute ) ? subRoute : null ;
    }

    public IEnumerable<SubRoute> GetAllSubRoutes()
    {
      return _ep2SubRoute.Values.Distinct() ;
    }


    #region Inner classes

    private class AutoRoutingCondition : ICommonRoutingCondition
    {
      private readonly SubRoute _subRoute ;

      public AutoRoutingCondition( Document document, SubRoute subRoute, int priority, double? forcedFixedHeight )
      {
        var documentData = DocumentMapper.Get( document ) ;

        _subRoute = subRoute ;
        Priority = priority ;
        IsRoutingOnPipeRacks = ( 0 < documentData.RackCollection.RackCount ) && subRoute.IsRoutingOnPipeSpace ;
        AllowHorizontalBranches = documentData.AllowHorizontalBranches( subRoute ) ;
        FixedBopHeight = forcedFixedHeight ;
      }

      public bool IsRoutingOnPipeRacks { get ; }
      public bool IsCrossingPipeRacks => false ;
      public bool IsRouteMergeEnabled => true ;
      public LineType Type => _subRoute.Route.ServiceType ;
      public int Priority { get ; }
      public LoopType LoopType => LoopType.Non ;

      public bool AllowHorizontalBranches { get ; }
      public double? FixedBopHeight { get ; set ; }
    }

    private class AutoRoutingSpatialConstraints : IAutoRoutingSpatialConstraints
    {
      public AutoRoutingSpatialConstraints( IEnumerable<IAutoRoutingEndPoint> fromEndPoints, IEnumerable<IAutoRoutingEndPoint> toEndPoints )
      {
        Starts = fromEndPoints ;
        Destination = toEndPoints ;
      }

      public IEnumerable<IAutoRoutingEndPoint> Starts { get ; }

      public IEnumerable<IAutoRoutingEndPoint> Destination { get ; }
    }

    #endregion
  }
}