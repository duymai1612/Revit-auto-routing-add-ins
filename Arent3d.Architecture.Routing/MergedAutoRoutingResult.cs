using System ;
using System.Collections.Generic ;
using System.Diagnostics ;
using System.IO ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Routing ;
using Arent3d.Utility ;
using MathLib ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Wrapper class of auto routing result.
  /// </summary>
  public class MergedAutoRoutingResult
  {
    private readonly IReadOnlyList<IAutoRoutingResult> _results ;

    public IReadOnlyCollection<IRouteEdge> RouteEdges { get ; }
    public IReadOnlyCollection<IRouteVertex> RouteVertices { get ; }

    private readonly IReadOnlyDictionary<IRouteEdge, PassingEndPointInfo> _passingEndPointInfo ;

    internal MergedAutoRoutingResult( IReadOnlyList<IAutoRoutingResult> results )
    {
      _results = results ;

      var routeVertices = new List<IRouteVertex>() ;
      var dummyBreakEndPoints = new Dictionary<int, IRouteVertex>() ;
      foreach ( var vertex in results.SelectMany( result => result.RouteVertices ) ) {
        if ( GetDummyBreakEndPoint( vertex ) is {} dbp ) {
          if ( dummyBreakEndPoints.ContainsKey( dbp.Id ) ) continue ;

          var breakPoint = new BreakPoint( vertex ) ;
          dummyBreakEndPoints.Add( dbp.Id, breakPoint ) ;
          routeVertices.Add( breakPoint ) ;
        }
        else {
          routeVertices.Add( vertex ) ;
        }
      }

      // IAutoRoutingResult.RouteEdges returns different instances between calls. MergedAutoRoutingResult will preserve them.
      var routeEdges = new List<IRouteEdge>() ;
      foreach ( var edge in results.SelectMany( result => result.RouteEdges ) ) {
        var startVertex = edge.Start ;
        if ( GetDummyBreakEndPoint( startVertex ) is { } dbp1 ) {
          startVertex = dummyBreakEndPoints[ dbp1.Id ] ;
        }

        var endVertex = edge.End ;
        if ( GetDummyBreakEndPoint( endVertex ) is { } dbp2 ) {
          endVertex = dummyBreakEndPoints[ dbp2.Id ] ;
        }

        routeEdges.Add( new RouteEdge( edge, startVertex, endVertex ) ) ;
      }

      RouteVertices = routeVertices ;
      RouteEdges = routeEdges ;
      _passingEndPointInfo = PassingEndPointInfo.CollectPassingEndPointInfo( RouteEdges ) ;
    }

    private static DummyBreakEndPoint? GetDummyBreakEndPoint( IRouteVertex vertex )
    {
      return ( vertex as TerminalPoint )?.LineInfo.GetEndPoint() as DummyBreakEndPoint ;
    }

    public PassingEndPointInfo GetPassingEndPointInfo( IRouteEdge edge )
    {
      if ( false == _passingEndPointInfo.TryGetValue( edge, out var info ) ) throw new ArgumentException() ;

      return info ;
    }

    [Conditional( "DEBUG" )]
    public void DebugExport( string fileName )
    {
      if ( 1 == _results.Count ) {
        _results[ 0 ].DebugExport( fileName ) ;
      }
      else {
        var dir = Path.GetDirectoryName( fileName ) ?? string.Empty ;
        var file = Path.GetFileNameWithoutExtension( fileName ) ?? string.Empty ;
        var ext = Path.GetExtension( fileName ) ?? string.Empty ;
        _results.ForEach( ( result, index ) => result.DebugExport( Path.Combine( dir, $"{file}_{index}{ext}" ) ) ) ;
      }
    }

    private class BreakPoint : IRouteVertex
    {
      private readonly IRouteVertex _baseVertex ;
      public Vector3d Position => _baseVertex.Position ;
      public IPipeDiameter PipeDiameter => _baseVertex.PipeDiameter ;
      public IAutoRoutingEndPoint LineInfo => _baseVertex.LineInfo ;
      public bool IsOverflow => _baseVertex.IsOverflow ;

      public BreakPoint( IRouteVertex vertex )
      {
        _baseVertex = vertex ;
      }
    }

    private class RouteEdge : IRouteEdge
    {
      private readonly IRouteEdge _baseEdge ;
      public IRouteVertex Start { get ; }
      public IRouteVertex End { get ; }
      public IAutoRoutingEndPoint LineInfo => _baseEdge.LineInfo ;
      public ILayerProperty RelatedLayer => _baseEdge.RelatedLayer ;
      public bool IsOverflowed => _baseEdge.IsOverflowed ;

      public RouteEdge( IRouteEdge edge, IRouteVertex startVertex, IRouteVertex endVertex )
      {
        _baseEdge = edge ;
        Start = startVertex ;
        End = endVertex ;
      }
    }
  }
}