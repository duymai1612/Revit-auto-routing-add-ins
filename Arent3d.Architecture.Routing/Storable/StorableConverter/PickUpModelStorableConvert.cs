using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( PickUpModel ) )]
  public class PickUpModelStorableConvert : StorableConverterBase<PickUpModel>
  {
    private enum SerializeField
    {
      Item,
      Floor,
      ConstructionItems,
      EquipmentType,
      ProductName,
      Use,
      UsageName,
      Construction,
      ModelNumber,
      Specification,
      Specification2,
      Size,
      Quantity,
      Tani,
      Supplement,
      Supplement2,
      Group,
      Layer,
      Classification,
      Standard,
      PickUpNumber,
      Direction,
      ProductCode
    }

    protected override PickUpModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var item = deserializer.GetString( SerializeField.Item ) ;
      var floor = deserializer.GetString( SerializeField.Floor ) ;
      var constructionItems = deserializer.GetString( SerializeField.ConstructionItems ) ;
      var equipmentType = deserializer.GetString( SerializeField.EquipmentType ) ;
      var productName = deserializer.GetString( SerializeField.ProductName ) ;
      var use = deserializer.GetString( SerializeField.Use ) ;
      var usageName = deserializer.GetString( SerializeField.UsageName ) ;
      var construction = deserializer.GetString( SerializeField.Construction ) ;
      var modelNumber = deserializer.GetString( SerializeField.ModelNumber ) ;
      var specification = deserializer.GetString( SerializeField.Specification ) ;
      var specification2 = deserializer.GetString( SerializeField.Specification2 ) ;
      var size = deserializer.GetString( SerializeField.Size ) ;
      var quantity = deserializer.GetString( SerializeField.Quantity ) ;
      var tani = deserializer.GetString( SerializeField.Tani ) ;
      var supplement = deserializer.GetString( SerializeField.Supplement ) ;
      var supplement2 = deserializer.GetString( SerializeField.Supplement2 ) ;
      var group = deserializer.GetString( SerializeField.Group ) ;
      var layer = deserializer.GetString( SerializeField.Layer ) ;
      var classification = deserializer.GetString( SerializeField.Classification ) ;
      var standard = deserializer.GetString( SerializeField.Standard ) ;
      var pickUpNumber = deserializer.GetString( SerializeField.PickUpNumber ) ;
      var direction = deserializer.GetString( SerializeField.Direction ) ;
      var productCode = deserializer.GetString( SerializeField.ProductCode ) ;

      return new PickUpModel( item, floor, constructionItems, equipmentType, productName, use, usageName, construction, modelNumber, specification, specification2, size, quantity, tani, supplement, supplement2, group, layer, classification, standard, pickUpNumber, direction, productCode ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, PickUpModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.Item, customTypeValue.Item ) ;
      serializerObject.AddNonNull( SerializeField.Floor, customTypeValue.Floor ) ;
      serializerObject.AddNonNull( SerializeField.ConstructionItems, customTypeValue.ConstructionItems ) ;
      serializerObject.AddNonNull( SerializeField.EquipmentType, customTypeValue.EquipmentType ) ;
      serializerObject.AddNonNull( SerializeField.ProductName, customTypeValue.ProductName ) ;
      serializerObject.AddNonNull( SerializeField.Use, customTypeValue.Use ) ;
      serializerObject.AddNonNull( SerializeField.Construction, customTypeValue.Construction ) ;
      serializerObject.AddNonNull( SerializeField.ModelNumber, customTypeValue.ModelNumber ) ;
      serializerObject.AddNonNull( SerializeField.Specification, customTypeValue.Specification ) ;
      serializerObject.AddNonNull( SerializeField.Specification2, customTypeValue.Specification2 ) ;
      serializerObject.AddNonNull( SerializeField.Size, customTypeValue.Size ) ;
      serializerObject.AddNonNull( SerializeField.Quantity, customTypeValue.Quantity ) ;
      serializerObject.AddNonNull( SerializeField.Tani, customTypeValue.Tani ) ;
      serializerObject.AddNonNull( SerializeField.Supplement, customTypeValue.Supplement ) ;
      serializerObject.AddNonNull( SerializeField.Supplement2, customTypeValue.Supplement2 ) ;
      serializerObject.AddNonNull( SerializeField.Group, customTypeValue.Group ) ;
      serializerObject.AddNonNull( SerializeField.Layer, customTypeValue.Layer ) ;
      serializerObject.AddNonNull( SerializeField.Classification, customTypeValue.Classification ) ;
      serializerObject.AddNonNull( SerializeField.Standard, customTypeValue.Standard ) ;
      serializerObject.AddNonNull( SerializeField.PickUpNumber, customTypeValue.PickUpNumber ) ;
      serializerObject.AddNonNull( SerializeField.Direction, customTypeValue.Direction ) ;
      serializerObject.AddNonNull( SerializeField.ProductCode, customTypeValue.ProductCode ) ;

      return serializerObject ;
    }
  }
}