using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( HiroiSetMasterModel ) )]
  public class HiroiSetMasterModelStorableConverter : StorableConverterBase<HiroiSetMasterModel>
  {
    private enum SerializeField
    {
      ParentPartModelNumber,
      ParentPartName,
      ParentPartsQuantity,
      MaterialCode1,
      Name1,
      Quantity1,
      MaterialCode2,
      Name2,
      Quantity2,
      MaterialCode3,
      Name3,
      Quantity3,
      MaterialCode4,
      Name4,
      Quantity4,
      MaterialCode5,
      Name5,
      Quantity5,
      MaterialCode6,
      Name6,
      Quantity6,
      MaterialCode7,
      Name7,
      Quantity7,
      MaterialCode8,
      Name8,
      Quantity8
    }

    protected override HiroiSetMasterModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var parentPartModelNumber = deserializer.GetString( SerializeField.ParentPartModelNumber ) ;
      var parentPartName = deserializer.GetString( SerializeField.ParentPartName ) ;
      var parentPartsQuantity = deserializer.GetString( SerializeField.ParentPartsQuantity ) ;
      var materialCode1 = deserializer.GetString( SerializeField.MaterialCode1 ) ;
      var name1 = deserializer.GetString( SerializeField.Name1 ) ;
      var quantity1 = deserializer.GetString( SerializeField.Quantity1 ) ;
      var materialCode2 = deserializer.GetString( SerializeField.MaterialCode2 ) ;
      var name2 = deserializer.GetString( SerializeField.Name2 ) ;
      var quantity2 = deserializer.GetString( SerializeField.Quantity2 ) ;
      var materialCode3 = deserializer.GetString( SerializeField.MaterialCode3 ) ;
      var name3 = deserializer.GetString( SerializeField.Name3 ) ;
      var quantity3 = deserializer.GetString( SerializeField.Quantity3 ) ;
      var materialCode4 = deserializer.GetString( SerializeField.MaterialCode4 ) ;
      var name4 = deserializer.GetString( SerializeField.Name4 ) ;
      var quantity4 = deserializer.GetString( SerializeField.Quantity4 ) ;
      var materialCode5 = deserializer.GetString( SerializeField.MaterialCode5 ) ;
      var name5 = deserializer.GetString( SerializeField.Name5 ) ;
      var quantity5 = deserializer.GetString( SerializeField.Quantity5 ) ;
      var materialCode6 = deserializer.GetString( SerializeField.MaterialCode6 ) ;
      var name6 = deserializer.GetString( SerializeField.Name6 ) ;
      var quantity6 = deserializer.GetString( SerializeField.Quantity6 ) ;
      var materialCode7 = deserializer.GetString( SerializeField.MaterialCode7 ) ;
      var name7 = deserializer.GetString( SerializeField.Name7 ) ;
      var quantity7 = deserializer.GetString( SerializeField.Quantity7 ) ;
      var materialCode8 = deserializer.GetString( SerializeField.MaterialCode8 ) ;
      var name8 = deserializer.GetString( SerializeField.Name8 ) ;
      var quantity8 = deserializer.GetString( SerializeField.Quantity8 ) ;

      return new HiroiSetMasterModel( parentPartModelNumber, parentPartName, parentPartsQuantity, materialCode1, name1, quantity1, materialCode2, name2, quantity2, materialCode3, name3, quantity3, materialCode4, name4, quantity4, materialCode5, name5, quantity5, materialCode6, name6, quantity6, materialCode7, name7, quantity7, materialCode8, name8, quantity8 ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, HiroiSetMasterModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.ParentPartModelNumber, customTypeValue.ParentPartModelNumber ) ;
      serializerObject.AddNonNull( SerializeField.ParentPartName, customTypeValue.ParentPartName ) ;
      serializerObject.AddNonNull( SerializeField.ParentPartsQuantity, customTypeValue.ParentPartsQuantity ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode1, customTypeValue.MaterialCode1 ) ;
      serializerObject.AddNonNull( SerializeField.Name1, customTypeValue.Name1 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity1, customTypeValue.Quantity1 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode2, customTypeValue.MaterialCode2 ) ;
      serializerObject.AddNonNull( SerializeField.Name2, customTypeValue.Name2 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity2, customTypeValue.Quantity2 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode3, customTypeValue.MaterialCode3 ) ;
      serializerObject.AddNonNull( SerializeField.Name3, customTypeValue.Name3 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity3, customTypeValue.Quantity3 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode4, customTypeValue.MaterialCode4 ) ;
      serializerObject.AddNonNull( SerializeField.Name4, customTypeValue.Name4 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity4, customTypeValue.Quantity4 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode5, customTypeValue.MaterialCode5 ) ;
      serializerObject.AddNonNull( SerializeField.Name5, customTypeValue.Name5 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity5, customTypeValue.Quantity5 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode6, customTypeValue.MaterialCode6 ) ;
      serializerObject.AddNonNull( SerializeField.Name6, customTypeValue.Name6 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity6, customTypeValue.Quantity6 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode7, customTypeValue.MaterialCode7 ) ;
      serializerObject.AddNonNull( SerializeField.Name7, customTypeValue.Name7 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity7, customTypeValue.Quantity7 ) ;
      serializerObject.AddNonNull( SerializeField.MaterialCode8, customTypeValue.MaterialCode8 ) ;
      serializerObject.AddNonNull( SerializeField.Name8, customTypeValue.Name8 ) ;
      serializerObject.AddNonNull( SerializeField.Quantity8, customTypeValue.Quantity8 ) ;

      return serializerObject ;
    }
  }
}