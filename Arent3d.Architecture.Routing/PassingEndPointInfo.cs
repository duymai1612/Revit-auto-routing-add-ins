using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Routing ;
using Arent3d.Utility ;

namespace Arent3d.Architecture.Routing
{
  public class PassingEndPointInfo
  {
    public static IReadOnlyDictionary<IRouteEdge, PassingEndPointInfo> CollectPassingEndPointInfo( IReadOnlyCollection<IRouteEdge> routeEdges )
    {
      var dic = new Dictionary<IRouteEdge, PassingEndPointInfo>() ;

      var linkInfo = new Dictionary<IRouteVertex, (List<IRouteEdge> Enter, List<IRouteEdge> Exit)>() ;
      foreach ( var edge in routeEdges ) {
        AddLinkInfo( linkInfo, edge.Start, edge, true ) ;
        AddLinkInfo( linkInfo, edge.End, edge, false ) ;
      }

      foreach ( var edge in routeEdges ) {
        SeekEndPoints( dic, linkInfo, edge, true ) ;
        SeekEndPoints( dic, linkInfo, edge, false ) ;
      }

      return dic ;
    }

    private static void AddLinkInfo( Dictionary<IRouteVertex, (List<IRouteEdge> Enter, List<IRouteEdge> Exit)> linkInfo, IRouteVertex vertex, IRouteEdge edge, bool isExit )
    {
      if ( false == linkInfo.TryGetValue( vertex, out var tuple ) ) {
        tuple = ( new List<IRouteEdge>(), new List<IRouteEdge>() ) ;
        linkInfo.Add( vertex, tuple ) ;
      }

      if ( isExit ) {
        tuple.Exit.Add( edge ) ;
      }
      else {
        tuple.Enter.Add( edge ) ;
      }
    }

    private static IEnumerable<KeyValuePair<EndPointKey, IEndPoint>> SeekEndPoints( Dictionary<IRouteEdge, PassingEndPointInfo> dic, IReadOnlyDictionary<IRouteVertex, (List<IRouteEdge> Enter, List<IRouteEdge> Exit)> linkInfo, IRouteEdge edge, bool seekFrom )
    {
      if ( false == dic.TryGetValue( edge, out var fromTo ) ) {
        fromTo = new PassingEndPointInfo() ;
        dic.Add( edge, fromTo ) ;
      }

      if ( seekFrom ) {
        if ( 0 == fromTo._fromEndPoints.Count ) {
          if ( edge.Start is TerminalPoint tp ) {
            fromTo.RegisterFrom( tp.LineInfo.GetEndPoint() ) ;
          }
          fromTo.RegisterFrom( linkInfo[ edge.Start ].Enter.SelectMany( e => SeekEndPoints( dic, linkInfo, e, true ) ) ) ;
        }

        return fromTo._fromEndPoints ;
      }
      else {
        if ( 0 == fromTo._toEndPoints.Count ) {
          if ( edge.End is TerminalPoint tp ) {
            fromTo.RegisterTo( tp.LineInfo.GetEndPoint() ) ;
          }
          fromTo.RegisterTo( linkInfo[ edge.End ].Exit.SelectMany( e => SeekEndPoints( dic, linkInfo, e, false ) ) ) ;
        }

        return fromTo._toEndPoints ;
      }
    }

    private void RegisterFrom( IEndPoint? endPoint )
    {
      if ( null != endPoint ) {
        _fromEndPoints.Add( endPoint.Key, endPoint ) ;
      }
    }
    private void RegisterTo( IEndPoint? endPoint )
    {
      if ( null != endPoint ) {
        _toEndPoints.Add( endPoint.Key, endPoint ) ;
      }
    }

    private void RegisterFrom( IEnumerable<KeyValuePair<EndPointKey, IEndPoint>> endPoints )
    {
      foreach ( var (key, endPoint) in endPoints ) {
        if ( _fromEndPoints.ContainsKey( key ) ) continue ;

        _fromEndPoints.Add( key, endPoint ) ;
      }
    }
    private void RegisterTo( IEnumerable<KeyValuePair<EndPointKey, IEndPoint>> endPoints )
    {
      foreach ( var (key, endPoint) in endPoints ) {
        if ( _toEndPoints.ContainsKey( key ) ) continue ;

        _toEndPoints.Add( key, endPoint ) ;
      }
    }

    private readonly Dictionary<EndPointKey, IEndPoint> _fromEndPoints = new() ;

    private readonly Dictionary<EndPointKey, IEndPoint> _toEndPoints = new() ;

    private PassingEndPointInfo()
    {
    }

    public IEnumerable<IEndPoint> FromEndPoints => _fromEndPoints.Values ;
    public IEnumerable<IEndPoint> ToEndPoints => _toEndPoints.Values ;

    public bool TryGetFromEndPoint( EndPointKey key, out IEndPoint endPoint ) => _fromEndPoints.TryGetValue( key, out endPoint ) ;
    public bool TryGetToEndPoint( EndPointKey key, out IEndPoint endPoint ) => _toEndPoints.TryGetValue( key, out endPoint ) ;
    

    
    public static PassingEndPointInfo CreatePassingEndPointInfo( IEndPoint fromEndPoint, IEndPoint toEndPoint )
    {
      var result = new PassingEndPointInfo() ;

      result.RegisterFrom( fromEndPoint ) ;
      result.RegisterTo( toEndPoint ) ;

      return result ;
    }
  }
}