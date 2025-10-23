using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( ConduitsModel ) )]
  public class ConduitsModelStorableConverter : StorableConverterBase<ConduitsModel>
  {
    private enum SerializeField
    {
      PipingType,
      Size,
      InnerCrossSectionalArea,
      Name,
      Classification
    }

    protected override ConduitsModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var pipingType = deserializer.GetString( SerializeField.PipingType ) ;
      var size = deserializer.GetString( SerializeField.Size ) ;
      var innerCrossSectionalArea = deserializer.GetString( SerializeField.InnerCrossSectionalArea ) ;
      var name = deserializer.GetString( SerializeField.Name ) ;
      var classification = deserializer.GetString( SerializeField.Classification ) ;

      return new ConduitsModel( pipingType, size, innerCrossSectionalArea, name, classification ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, ConduitsModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.PipingType, customTypeValue.PipingType ) ;
      serializerObject.AddNonNull( SerializeField.Size, customTypeValue.Size ) ;
      serializerObject.AddNonNull( SerializeField.InnerCrossSectionalArea, customTypeValue.InnerCrossSectionalArea ) ;
      serializerObject.AddNonNull( SerializeField.Name, customTypeValue.Name ) ;
      serializerObject.AddNonNull( SerializeField.Classification, customTypeValue.Classification ) ;

      return serializerObject ;
    }
  }
}