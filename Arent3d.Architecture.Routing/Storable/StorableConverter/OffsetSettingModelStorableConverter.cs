using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( OffsetSettingModel ) )]
  internal class OffsetSettingModelStorableConverter : StorableConverterBase<OffsetSettingModel>
  {
    private enum SerializeField
    {
      Offset
    }

    protected override OffsetSettingModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var offset = deserializer.GetDouble( SerializeField.Offset ) ;

      return new OffsetSettingModel( offset ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, OffsetSettingModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.Add( SerializeField.Offset, customTypeValue.Offset ) ;

      return serializerObject ;
    }
  }
}