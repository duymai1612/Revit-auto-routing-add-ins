using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.StorableCaches
{
  public class RouteCache : StorableCache<RouteCache, Route>
  {
    private RouteCache( Document document ) : base( document )
    {
    }

    protected override Route CreateNewStorable( Document document, string name ) => new Route( document, name ) ;

    public IEnumerable<SubRoute> CollectAllSubRoutes() => Values.SelectMany( route => route.SubRoutes ) ;

    public SubRoute? GetSubRoute( string routeName, int subRouteIndex )
    {
      if ( false == TryGetValue( routeName, out var route ) ) return null ;
      return route.GetSubRoute( subRouteIndex ) ;
    }
    public SubRoute? GetSubRoute( SubRouteInfo subRouteInfo ) => GetSubRoute( subRouteInfo.RouteName, subRouteInfo.SubRouteIndex ) ;
  }
}