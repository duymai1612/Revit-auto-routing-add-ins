using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( DetailSymbolModel ) )]
  public class DetailSymbolModelStorableConvert : StorableConverterBase<DetailSymbolModel>
  {
    private enum SerializeField
    {
      DetailSymbolId,
      DetailSymbol,
      ConduitId,
      RouteName,
      Code,
      LineIds,
      IsParentSymbol
    }

    protected override DetailSymbolModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var detailSymbolId = deserializer.GetString( SerializeField.DetailSymbolId ) ;
      var detailSymbol = deserializer.GetString( SerializeField.DetailSymbol ) ;
      var conduitId = deserializer.GetString( SerializeField.ConduitId ) ;
      var routeName = deserializer.GetString( SerializeField.RouteName ) ;
      var code = deserializer.GetString( SerializeField.Code ) ;
      var lineIds = deserializer.GetString( SerializeField.LineIds ) ;
      var isParentSymbol = deserializer.GetBool( SerializeField.IsParentSymbol ) ;

      return new DetailSymbolModel( detailSymbolId, detailSymbol, conduitId, routeName, code, lineIds, isParentSymbol ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, DetailSymbolModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.DetailSymbolId, customTypeValue.DetailSymbolId ) ;
      serializerObject.AddNonNull( SerializeField.DetailSymbol, customTypeValue.DetailSymbol ) ;
      serializerObject.AddNonNull( SerializeField.ConduitId, customTypeValue.ConduitId ) ;
      serializerObject.AddNonNull( SerializeField.RouteName, customTypeValue.RouteName ) ;
      serializerObject.AddNonNull( SerializeField.Code, customTypeValue.Code ) ;
      serializerObject.AddNonNull( SerializeField.LineIds, customTypeValue.LineIds ) ;
      serializerObject.Add( SerializeField.IsParentSymbol, customTypeValue.IsParentSymbol ) ;

      return serializerObject ;
    }
  }
}