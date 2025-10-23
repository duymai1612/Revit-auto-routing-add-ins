using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  [StorableConverterOf( typeof( WiresAndCablesModel ) )]
  public class WiresAndCablesModelStorableConverter : StorableConverterBase<WiresAndCablesModel>
  {
    private enum SerializeField
    {
      WireType,
      DiameterOrNominal,
      DOrA,
      NumberOfHeartsOrLogarithm,
      COrP,
      CrossSectionalArea,
      Name,
      Classification,
      FinishedOuterDiameter,
      NumberOfConnections
    }

    protected override WiresAndCablesModel Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var wireType = deserializer.GetString( SerializeField.WireType ) ;
      var diameterOrNominal = deserializer.GetString( SerializeField.DiameterOrNominal ) ;
      var dOrA = deserializer.GetString( SerializeField.DOrA ) ;
      var numberOfHeartsOrLogarithm = deserializer.GetString( SerializeField.NumberOfHeartsOrLogarithm ) ;
      var cOrP = deserializer.GetString( SerializeField.COrP ) ;
      var crossSectionalArea = deserializer.GetString( SerializeField.CrossSectionalArea ) ;
      var name = deserializer.GetString( SerializeField.Name ) ;
      var classification = deserializer.GetString( SerializeField.Classification ) ;
      var finishedOuterDiameter = deserializer.GetString( SerializeField.FinishedOuterDiameter ) ;
      var numberOfConnections = deserializer.GetString( SerializeField.NumberOfConnections ) ;

      return new WiresAndCablesModel( wireType, diameterOrNominal, dOrA, numberOfHeartsOrLogarithm, cOrP, crossSectionalArea, name, classification, finishedOuterDiameter, numberOfConnections ) ;
    }

    protected override ISerializerObject Serialize( Element storedElement, WiresAndCablesModel customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.AddNonNull( SerializeField.WireType, customTypeValue.WireType ) ;
      serializerObject.AddNonNull( SerializeField.DiameterOrNominal, customTypeValue.DiameterOrNominal ) ;
      serializerObject.AddNonNull( SerializeField.DOrA, customTypeValue.DOrA ) ;
      serializerObject.AddNonNull( SerializeField.NumberOfHeartsOrLogarithm, customTypeValue.NumberOfHeartsOrLogarithm ) ;
      serializerObject.AddNonNull( SerializeField.COrP, customTypeValue.COrP ) ;
      serializerObject.AddNonNull( SerializeField.CrossSectionalArea, customTypeValue.CrossSectionalArea ) ;
      serializerObject.AddNonNull( SerializeField.Name, customTypeValue.Name ) ;
      serializerObject.AddNonNull( SerializeField.Classification, customTypeValue.Classification ) ;
      serializerObject.AddNonNull( SerializeField.FinishedOuterDiameter, customTypeValue.FinishedOuterDiameter ) ;
      serializerObject.AddNonNull( SerializeField.NumberOfConnections, customTypeValue.NumberOfConnections ) ;

      return serializerObject ;
    }
  }
}