using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.FittingSizeCalculators ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public abstract class AutoRoutingTargetGenerator
  {
    protected Document Document { get ; }
    private readonly Dictionary<(string, int), List<SubRoute>> _groups = new() ;

    protected AutoRoutingTargetGenerator( Document document )
    {
      Document = document ;
    }

    public IEnumerable<IReadOnlyCollection<AutoRoutingTarget>> Create( IReadOnlyCollection<Route> routes, IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> routeConditionDictionary )
    {
      var priorities = CollectPriorities( routes ) ;

      foreach ( var route in routes.OrderBy( r => priorities[ r ] ) ) {
        foreach ( var subRoute in route.SubRoutes ) {
          AddSubRoute( subRoute ) ;
        }
      }

      return _groups.Values.Distinct().Select( list => GenerateAutoRoutingTarget( list, priorities, routeConditionDictionary ) ) ;
    }

    private void AddSubRoute( SubRoute subRoute )
    {
      List<List<SubRoute>>? relatedSubRoutes = null ;

      foreach ( var routeEndPoint in subRoute.AllEndPoints.OfType<RouteEndPoint>() ) {
        if ( false == _groups.TryGetValue( ( routeEndPoint.RouteName, routeEndPoint.SubRouteIndex ), out var list ) ) continue ;

        relatedSubRoutes ??= new List<List<SubRoute>>() ;
        relatedSubRoutes.Add( list ) ;
      }

      if ( null == relatedSubRoutes ) {
        _groups.Add( ( subRoute.Route.RouteName, subRoute.SubRouteIndex ), new List<SubRoute> { subRoute } ) ;
      }
      else {
        var mergedList = relatedSubRoutes[ 0 ] ;
        if ( 1 < relatedSubRoutes.Count ) {
          foreach ( var list in relatedSubRoutes.Skip( 1 ) ) {
            mergedList.AddRange( list ) ;
            foreach ( var sr in list ) {
              _groups[ ( sr.Route.RouteName, sr.SubRouteIndex ) ] = mergedList ;
            }
          }
        }

        mergedList.Add( subRoute ) ;
        _groups.Add( ( subRoute.Route.RouteName, subRoute.SubRouteIndex ), mergedList ) ;
      }
    }

    protected abstract IReadOnlyCollection<AutoRoutingTarget> GenerateAutoRoutingTarget( IReadOnlyCollection<SubRoute> subRoutes, IReadOnlyDictionary<Route, int> priorities, IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> routeConditionDictionary ) ;

    private static IReadOnlyDictionary<Route, int> CollectPriorities( IReadOnlyCollection<Route> routes )
    {
      var dic = new Dictionary<Route, int>() ;

      var index = 0 ;
      var routesToRemove = new List<Route>() ;
      var routesToParents = routes.ToDictionary( route => route, route => route.GetParentBranches() ) ;

      while ( 0 < routesToParents.Count ) {
        routesToRemove.Clear() ;
        foreach ( var (route, parents) in routesToParents ) {
          if ( 0 == parents.Count( routesToParents.ContainsKey ) ) {
            dic.Add( route, index ) ;
            routesToRemove.Add( route ) ;
          }
        }

        if ( routesToParents.Count == routesToRemove.Count ) break ;

        // next layers.
        foreach ( var route in routesToRemove ) {
          routesToParents.Remove( route ) ;
        }

        foreach ( var set in routesToParents.Values ) {
          set.ExceptWith( routesToRemove ) ;
        }

        ++index ;
      }

      if ( dic.Count != routes.Count ) throw new InvalidOperationException() ;

      return dic ;
    }
  }
}