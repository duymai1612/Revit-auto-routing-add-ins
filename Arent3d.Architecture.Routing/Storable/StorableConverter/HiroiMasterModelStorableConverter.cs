using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( HiroiMasterModel ) )]
  public class HiroiMasterModelStorableConverter : StorableConverterBase<HiroiMasterModel>
  {
    private enum SerializeField
    {
      Setubisyu,
      Syurui,
      Buzaicd,
      Hinmei,
      Kikaku,
      Tani,
      Buzaisyu,
      Hinmeicd,
      Ryakumeicd,
      Type,
      Size1,
      Size2
    }

    protected override HiroiMasterModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var setubisyu = deserializer.GetString( SerializeField.Setubisyu ) ;
      var syurui = deserializer.GetString( SerializeField.Syurui ) ;
      var buzaicd = deserializer.GetString( SerializeField.Buzaicd ) ;
      var hinmei = deserializer.GetString( SerializeField.Hinmei ) ;
      var kikaku = deserializer.GetString( SerializeField.Kikaku ) ;
      var tani = deserializer.GetString( SerializeField.Tani ) ;
      var buzaisyu = deserializer.GetString( SerializeField.Buzaisyu ) ;
      var hinmeicd = deserializer.GetString( SerializeField.Hinmeicd ) ;
      var ryakumeicd = deserializer.GetString( SerializeField.Ryakumeicd ) ;
      var type = deserializer.GetString( SerializeField.Type ) ;
      var size1 = deserializer.GetString( SerializeField.Size1 ) ;
      var size2 = deserializer.GetString( SerializeField.Size2 ) ;

      return new HiroiMasterModel( setubisyu, syurui, buzaicd, hinmei, kikaku, tani, buzaisyu, hinmeicd, ryakumeicd, type, size1, size2 ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, HiroiMasterModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.Setubisyu, customTypeValue.Setubisyu ) ;
      serializerObject.AddNonNull( SerializeField.Syurui, customTypeValue.Syurui ) ;
      serializerObject.AddNonNull( SerializeField.Buzaicd, customTypeValue.Buzaicd ) ;
      serializerObject.AddNonNull( SerializeField.Hinmei, customTypeValue.Hinmei ) ;
      serializerObject.AddNonNull( SerializeField.Kikaku, customTypeValue.Kikaku ) ;
      serializerObject.AddNonNull( SerializeField.Tani, customTypeValue.Tani ) ;
      serializerObject.AddNonNull( SerializeField.Buzaisyu, customTypeValue.Buzaisyu ) ;
      serializerObject.AddNonNull( SerializeField.Hinmeicd, customTypeValue.Hinmeicd ) ;
      serializerObject.AddNonNull( SerializeField.Ryakumeicd, customTypeValue.Ryakumeicd ) ;
      serializerObject.AddNonNull( SerializeField.Type, customTypeValue.Type ) ;
      serializerObject.AddNonNull( SerializeField.Size1, customTypeValue.Size1 ) ;
      serializerObject.AddNonNull( SerializeField.Size2, customTypeValue.Size2 ) ;

      return serializerObject ;
    }
  }
}