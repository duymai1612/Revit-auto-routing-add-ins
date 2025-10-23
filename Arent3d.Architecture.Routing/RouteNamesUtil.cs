using System ;
using System.Collections.Generic ;
using Arent3d.Utility.Serialization ;

namespace Arent3d.Architecture.Routing
{
  public static class RouteNamesUtil
  {
    private enum ListSerializeField
    {
      RouteName
    }

    public static string StringifyRouteNames( IEnumerable<string> routeNames )
    {
      var serializer = new SerializerObject<ListSerializeField>() ;
      serializer.AddNonNull( ListSerializeField.RouteName, routeNames ) ;
      return serializer.ToString() ;
    }

    public static IReadOnlyCollection<string> ParseRouteNames( string str )
    {
      var deserializer = new DeserializerObject<ListSerializeField>( str ) ;
      var array = deserializer.GetNonNullStringArray( ListSerializeField.RouteName ) ;
      if ( array == null ) return Array.Empty<string>() ;
      return array ;
    }
  }
}