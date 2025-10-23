using System ;
using Arent3d.Utility.Serialization ;

namespace Arent3d.Architecture.Routing
{
  public class SubRouteInfo : ISerializable, IEquatable<SubRouteInfo>
  {
    public string RouteName { get ; private set ; }
    public int SubRouteIndex { get ; private set ; }

    internal void ReplaceRouteName( string newRouteName )
    {
      RouteName = newRouteName ;
    }

    internal SubRouteInfo( string routeName, int subRouteIndex )
    {
      RouteName = routeName ;
      SubRouteIndex = subRouteIndex ;
    }

    public SubRouteInfo( SubRoute subRoute ) : this( subRoute.Route.RouteName, subRoute.SubRouteIndex )
    {
    }

    public void Deconstruct( out string routeName, out int subRouteIndex )
    {
      routeName = RouteName ;
      subRouteIndex = SubRouteIndex ;
    }


    #region ISerializable

    public static SubRouteInfo CreateForDeserialize() => new SubRouteInfo( string.Empty, -1 ) ;

    private enum SerializeField
    {
      RouteName,
      SubRouteIndex,
    }

    void ISerializable.SerializeInto( SerializerObject serializer )
    {
      var serializerObject = serializer.Of<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.RouteName, RouteName ) ;
      serializerObject.Add( SerializeField.SubRouteIndex, SubRouteIndex ) ;
    }

    bool ISerializable.DeserializeFrom( DeserializerObject deserializer )
    {
      var deserializerObject = deserializer.Of<SerializeField>() ;

      if ( deserializerObject.GetString( SerializeField.RouteName ) is not { } routeName ) return false ;
      if ( deserializerObject.GetInt( SerializeField.SubRouteIndex ) is not { } subRouteIndex ) return false ;

      RouteName = routeName ;
      SubRouteIndex = subRouteIndex ;
      
      return true ;
    }

    #endregion

    #region IEquatable
    
    public bool Equals( SubRouteInfo? other )
    {
      if ( ReferenceEquals( null, other ) ) return false ;
      if ( ReferenceEquals( this, other ) ) return true ;
      return RouteName == other.RouteName && SubRouteIndex == other.SubRouteIndex ;
    }

    public override bool Equals( object? obj )
    {
      if ( ReferenceEquals( null, obj ) ) return false ;
      if ( ReferenceEquals( this, obj ) ) return true ;
      if ( obj.GetType() != this.GetType() ) return false ;
      return Equals( (SubRouteInfo)obj ) ;
    }

    public override int GetHashCode()
    {
      unchecked {
        return ( RouteName.GetHashCode() * 397 ) ^ SubRouteIndex ;
      }
    }

    public static bool operator ==( SubRouteInfo? left, SubRouteInfo? right )
    {
      return Equals( left, right ) ;
    }

    public static bool operator !=( SubRouteInfo? left, SubRouteInfo? right )
    {
      return ! Equals( left, right ) ;
    }

    #endregion
  }
}