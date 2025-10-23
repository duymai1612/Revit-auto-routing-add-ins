using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase
{
  /// <summary>
  /// Detects <see cref="RouteSegment"/>s which passes through an element.
  /// </summary>
  public class RouteSegmentDetector
  {
    public Document Document { get ; }
    public string RouteName { get ; }
    public int SubRouteIndex { get ; }
    private readonly HashSet<EndPointKey> _fromElms = new() ;
    private readonly HashSet<EndPointKey> _toElms = new() ;

    /// <summary>
    /// Create a <see cref="RouteSegmentDetector"/>.
    /// </summary>
    /// <param name="subRoute">A <see cref="SubRoute"/> which can be affected by the passed-through element.</param>
    /// <param name="elementToPassThrough">A passed-through element.</param>
    public RouteSegmentDetector( SubRoute subRoute, Element elementToPassThrough )
    {
      Document = subRoute.Route.Document ;
      RouteName = subRoute.Route.RouteName ;
      SubRouteIndex = subRoute.SubRouteIndex ;

      CollectEndPoints( elementToPassThrough, true, _fromElms ) ;
      CollectEndPoints( elementToPassThrough, false, _toElms ) ;
    }

    private static void CollectEndPoints( Element element, bool isFrom, HashSet<EndPointKey> foundElms )
    {
      foundElms.UnionWith( element.GetNearestEndPoints( isFrom ).Select( ep => ep.Key ) ) ;
    }

    /// <summary>
    /// Returns a pass point index which is after the pass-through element
    /// </summary>
    /// <param name="info">Route info.</param>
    /// <returns>
    /// <para>Pass point index.</para>
    /// <para>0: The passed-through element is between the from-side connector and the first pass point (when no pass points, the to-side connector).</para>
    /// <para>1 to (info.PassPoints.Length - 1): The passed-through element is between the (k-1)-th pass point and the k-th pass point.</para>
    /// <para>info.PassPoints.Length: The passed-through element is between the last pass point and the to-side connector.</para>
    /// <para>-1: Not passed through.</para>
    /// </returns>
    public bool IsPassingThrough( RouteSegment info )
    {
      return ContainsFromEndPoint( info.FromEndPoint ) && ContainsToEndPoint( info.ToEndPoint ) ;
    }

    private bool ContainsFromEndPoint( IEndPoint endPoint )
    {
      if ( endPoint is RouteEndPoint routeEndPoint ) {
        if ( routeEndPoint.ParentSubRoute() is not { } subRoute ) return false ;
        return subRoute.Segments.Any( seg => ContainsFromEndPoint( seg.FromEndPoint ) ) ;
      }

      return _fromElms.Contains( endPoint.Key ) ;
    }

    private bool ContainsToEndPoint( IEndPoint endPoint )
    {
      if ( endPoint is RouteEndPoint routeEndPoint ) {
        if ( routeEndPoint.ParentSubRoute() is not { } subRoute ) return false ;
        return subRoute.Segments.Any( seg => ContainsToEndPoint( seg.ToEndPoint ) ) ;
      }

      return _toElms.Contains( endPoint.Key ) ;
    }
  }
}