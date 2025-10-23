using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( HiroiSetCdMasterModel ) )]
  public class HiroiSetCdMasterModelStorableConverter : StorableConverterBase<HiroiSetCdMasterModel>
  {
    private enum SerializeField
    {
      SetCode,
      QuantityParentPartModelNumber,
      LengthParentPartModelNumber,
      ConstructionClassification
    }

    protected override HiroiSetCdMasterModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var setCode = deserializer.GetString( SerializeField.SetCode ) ;
      var quantityParentPartModelNumber = deserializer.GetString( SerializeField.QuantityParentPartModelNumber ) ;
      var lengthParentPartModelNumber = deserializer.GetString( SerializeField.LengthParentPartModelNumber ) ;
      var constructionClassification = deserializer.GetString( SerializeField.ConstructionClassification ) ;

      return new HiroiSetCdMasterModel( setCode, quantityParentPartModelNumber, lengthParentPartModelNumber, constructionClassification ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, HiroiSetCdMasterModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.SetCode, customTypeValue.SetCode ) ;
      serializerObject.AddNonNull( SerializeField.QuantityParentPartModelNumber, customTypeValue.QuantityParentPartModelNumber ) ;
      serializerObject.AddNonNull( SerializeField.LengthParentPartModelNumber, customTypeValue.LengthParentPartModelNumber ) ;
      serializerObject.AddNonNull( SerializeField.ConstructionClassification, customTypeValue.ConstructionClassification ) ;

      return serializerObject ;
    }
  }
}