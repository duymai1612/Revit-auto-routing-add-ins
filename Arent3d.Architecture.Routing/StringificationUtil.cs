using System ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public static class StringificationUtil
  {
    #region EndPointKey

    private enum EndPointKeyFields
    {
      Type,
      Param,
    }

    public static void AddNonNull( this ISerializerObject serializerObject, string name, EndPointKey endPointKey )
    {
      serializerObject.Add( name, subSerializerObject => subSerializerObject.Of<EndPointKeyFields>().Add( endPointKey ) ) ;
    }

    public static void AddNonNull<TFieldEnum>( this SerializerObject<TFieldEnum> serializerObject, TFieldEnum field, EndPointKey endPointKey ) where TFieldEnum : struct, Enum
    {
      serializerObject.Add<EndPointKeyFields>( field, subSerializerObject => subSerializerObject.Add( endPointKey ) ) ;
    }

    public static void AddNullable( this ISerializerObject serializerObject, string name, EndPointKey? endPointKey )
    {
      if ( endPointKey is { } key ) {
        serializerObject.AddNonNull( name, key ) ;
      }
      else {
        serializerObject.AddNull( name ) ;
      }
    }

    public static void AddNullable<TFieldEnum>( this SerializerObject<TFieldEnum> serializerObject, TFieldEnum field, EndPointKey? endPointKey ) where TFieldEnum : struct, Enum
    {
      if ( endPointKey is { } key ) {
        serializerObject.AddNonNull( field, key ) ;
      }
      else {
        serializerObject.AddNull( field ) ;
      }
    }

    private static void Add( this SerializerObject<EndPointKeyFields> serializerObject, EndPointKey endPointKey )
    {
      serializerObject.AddNonNull( EndPointKeyFields.Type, endPointKey.Type ) ;
      serializerObject.AddNonNull( EndPointKeyFields.Param, endPointKey.Param ) ;
    }

    public static EndPointKey? GetEndPointKey( this IDeserializerObject deserializerObject, string name )
    {
      return deserializerObject.GetObject( name, subDeserializerObject => subDeserializerObject.Of<EndPointKeyFields>().GetEndPointKey() ) ;
    }

    public static EndPointKey? GetEndPointKey<TFieldEnum>( this DeserializerObject<TFieldEnum> deserializerObject, TFieldEnum field ) where TFieldEnum : struct, Enum
    {
      return deserializerObject.GetObject<EndPointKey, EndPointKeyFields>( field, subDeserializerObject => subDeserializerObject.GetEndPointKey() ) ;
    }

    private static EndPointKey? GetEndPointKey( this DeserializerObject<EndPointKeyFields> deserializerObject )
    {
      if ( deserializerObject.GetString( EndPointKeyFields.Type ) is not { } type ) return null ;
      if ( deserializerObject.GetString( EndPointKeyFields.Param ) is not { } param ) return null ;
      return new EndPointKey( type, param ) ;
    }

    #endregion

    #region XYZ

    private enum XYZFields
    {
      X,
      Y,
      Z,
    }

    public static void AddNonNull( this ISerializerObject serializerObject, string name, XYZ xyz )
    {
      serializerObject.Add( name, subSerializerObject => subSerializerObject.Of<XYZFields>().Add( xyz ) ) ;
    }

    public static void AddNonNull<TFieldEnum>( this SerializerObject<TFieldEnum> serializerObject, TFieldEnum field, XYZ xyz ) where TFieldEnum : struct, Enum
    {
      serializerObject.Add<XYZFields>( field, subSerializerObject => subSerializerObject.Add( xyz ) ) ;
    }

    private static void Add( this SerializerObject<XYZFields> serializerObject, XYZ xyz )
    {
      serializerObject.Add( XYZFields.X, xyz.X ) ;
      serializerObject.Add( XYZFields.Y, xyz.Y ) ;
      serializerObject.Add( XYZFields.Z, xyz.Z ) ;
    }

    public static void AddNullable( this ISerializerObject serializerObject, string name, XYZ? xyz )
    {
      if ( xyz is { } vec ) {
        serializerObject.AddNonNull( name, vec ) ;
      }
      else {
        serializerObject.AddNull( name ) ;
      }
    }

    public static void AddNullable<TFieldEnum>( this SerializerObject<TFieldEnum> serializerObject, TFieldEnum field, XYZ? xyz ) where TFieldEnum : struct, Enum
    {
      if ( xyz is { } vec ) {
        serializerObject.AddNonNull( field, vec ) ;
      }
      else {
        serializerObject.AddNull( field ) ;
      }
    }

    public static XYZ? GetXYZ( this IDeserializerObject deserializerObject, string name )
    {
      return deserializerObject.GetObject( name, subDeserializerObject => subDeserializerObject.Of<XYZFields>().GetXYZ() ) ;
    }

    public static XYZ? GetXYZ<TFieldEnum>( this DeserializerObject<TFieldEnum> deserializerObject, TFieldEnum field ) where TFieldEnum : struct, Enum
    {
      return deserializerObject.GetObject<XYZ, XYZFields>( field, subDeserializerObject => subDeserializerObject.GetXYZ() ) ;
    }

    private static XYZ? GetXYZ( this DeserializerObject<XYZFields> deserializerObject )
    {
      if ( deserializerObject.GetDouble( XYZFields.X ) is not { } x ) return null ;
      if ( deserializerObject.GetDouble( XYZFields.Y ) is not { } y ) return null ;
      if ( deserializerObject.GetDouble( XYZFields.Z ) is not { } z ) return null ;
      return new XYZ( x, y, z ) ;
    }

    #endregion
  }
}