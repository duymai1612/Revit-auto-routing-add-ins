using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase
{
  internal class EndPointDictionaryForImport
  {
    private readonly Document _document ;
    private readonly Dictionary<string, IEndPoint> _dic = new() ;

    public EndPointDictionaryForImport( Document document )
    {
      _document = document ;
    }

    public IEndPoint? GetEndPoint( string routeName, string key, IEndPoint? endPoint )
    {
      if ( string.IsNullOrEmpty( key ) ) {
        return endPoint ;
      }

      if ( _dic.TryGetValue( key, out var prevEndPoint ) ) {
        // use previous indicator.
        return prevEndPoint ;
      }

      if ( null == endPoint ) return null ;

      GenerateInstance( endPoint, routeName ) ;
      _dic.Add( key, endPoint ) ;

      return endPoint ;
    }

    private static void GenerateInstance( IEndPoint endPoint, string routeName )
    {
      endPoint.GenerateInstance( routeName ) ;
    }
  }

  internal class EndPointDictionaryForExport : IEndPointVisitor<(string Key, IEndPoint EndPoint)>
  {
    private int _index = 0 ;
    private readonly Document _document ;
    private readonly Dictionary<EndPointKey, string> _sameNameDic = new() ;

    public EndPointDictionaryForExport( Document document )
    {
      _document = document ;
    }

    public (string Key, IEndPoint EndPoint) GetEndPoint( IEndPoint endPoint )
    {
      var (key, ep) = ToExportingEndPoint( endPoint ) ;
      if ( string.IsNullOrEmpty( key ) ) {
        return ( string.Empty, ep ) ;
      }

      return ( key, ep ) ;
    }

    private (string Key, IEndPoint EndPoint) ToExportingEndPoint( IEndPoint endPoint )
    {
      return endPoint.Accept( this ) ;
    }

    public (string Key, IEndPoint EndPoint) Visit( ConnectorEndPoint endPoint )
    {
      return ( string.Empty, endPoint ) ;
    }

    public (string Key, IEndPoint EndPoint) Visit( PassPointEndPoint endPoint )
    {
      return ( RegisterEndPoint( endPoint ), endPoint ) ;
    }

    public (string Key, IEndPoint EndPoint) Visit( PassPointBranchEndPoint endPoint )
    {
      return ( RegisterEndPoint( endPoint ), endPoint ) ;
    }

    public (string Key, IEndPoint EndPoint) Visit( RouteEndPoint endPoint )
    {
      return (string.Empty, endPoint ) ;
    }

    public (string Key, IEndPoint EndPoint) Visit( TerminatePointEndPoint endPoint )
    {
      return (string.Empty, endPoint ) ;
    }

    private string RegisterEndPoint( IEndPoint endPoint )
    {
      if ( false == _sameNameDic.TryGetValue( endPoint.Key, out var key ) ) {
        key = $"pp#{++_index}" ;
        _sameNameDic.Add( endPoint.Key, key ) ;
      }

      return key ;
    }
  }
}