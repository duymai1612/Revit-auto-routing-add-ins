using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Arent3d.Routing ;
using Arent3d.Utility ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing.EndPoints
{
  public static class EndPointExtensions
  {
    public static XYZ ForEndPointType( this XYZ direction, bool isFrom ) => isFrom ? direction : -direction ;

    public static Vector3d ForEndPointType( this Vector3d direction, bool isFrom ) => isFrom ? direction : -direction ;

    public static IEndPoint? GetEndPoint( this IAutoRoutingEndPoint endPoint )
    {
      return endPoint switch
      {
        AutoRoutingEndPoint ep => ep.EndPoint,
        IPseudoEndPoint pep => pep.Source.GetEndPoint(),
        _ => null,
      } ;
    }

    public static Level? GetLevel( this IEndPoint endPoint, Document document )
    {
      var levelId = endPoint.GetLevelId( document ) ;
      if ( ElementId.InvalidElementId == levelId ) return null ;

      return document.GetElementById<Level>( levelId ) ;
    }



    private enum ListSerializeField
    {
      EndPoints,
    }
    private enum EndPointSerializeField
    {
      Type,
      Parameters,
    }

    public static string Stringify( this IEnumerable<IEndPoint> endPoints )
    {
      var serializer = new SerializerObject<ListSerializeField>() ;
      serializer.AddNonNull( ListSerializeField.EndPoints, endPoints.Select( Stringify ) ) ;
      return serializer.ToString() ;
    }

    public static string Stringify( this IEndPoint endPoint )
    {
      var serializer = new SerializerObject<EndPointSerializeField>() ;
      serializer.AddNonNull( EndPointSerializeField.Type, endPoint.TypeName ) ;
      serializer.AddNonNull( EndPointSerializeField.Parameters, endPoint.ParameterString ) ;
      return serializer.ToString() ;
    }

    public static IEnumerable<IEndPoint> ParseEndPoints( this Document document, string str )
    {
      var deserializer = new DeserializerObject<ListSerializeField>( str ) ;
      var array = deserializer.GetNonNullStringArray( ListSerializeField.EndPoints ) ;
      return array?.Select( document.ParseEndPoint ).NonNull() ?? Enumerable.Empty<IEndPoint>() ;
    }

    public static IEndPoint? ParseEndPoint( this Document document, string str )
    {
      var deserializer = new DeserializerObject<EndPointSerializeField>(str) ;
      if ( deserializer.GetString( EndPointSerializeField.Type ) is not { } endPontType ) return null ;
      if ( deserializer.GetString( EndPointSerializeField.Parameters ) is not { } paramString ) return null ;

      return document.ParseEndPoint( endPontType, paramString ) ;
    }

    public static IEndPoint? ParseEndPoint( this Document document, string endPointType, string parameters )
    {
      return endPointType switch
      {
        ConnectorEndPoint.Type => ConnectorEndPoint.ParseParameterString( document, parameters ),
        PassPointEndPoint.Type => PassPointEndPoint.ParseParameterString( document, parameters ),
        PassPointBranchEndPoint.Type => PassPointBranchEndPoint.ParseParameterString( document, parameters ),
        RouteEndPoint.Type => RouteEndPoint.ParseParameterString( document, parameters ),
        TerminatePointEndPoint.Type => TerminatePointEndPoint.ParseParameterString( document, parameters ),
        _ => null,
      } ;
    }

    public static IEndPoint? ToEndPointOverSubRoute( this IEndPoint endPoint, Document document )
    {
      if ( endPoint is not PassPointBranchEndPoint passPointBranchEndPoint ) return endPoint ;

      return document.ToEndPoint( passPointBranchEndPoint.EndPointKeyOverSubRoute ) ;
    }

    private static IEndPoint? ToEndPoint( this Document document, EndPointKey key )
    {
      return key.Type switch
      {
        ConnectorEndPoint.Type => ConnectorEndPoint.FromKeyParam( document, key.Param ),
        PassPointEndPoint.Type => PassPointEndPoint.FromKeyParam( document, key.Param ),
        PassPointBranchEndPoint.Type => null,
        RouteEndPoint.Type => RouteEndPoint.FromKeyParam( document, key.Param ),
        TerminatePointEndPoint.Type => TerminatePointEndPoint.FromKeyParam( document, key.Param ),
        _ => null,
      } ;
    }
  }
}