using System;
using Arent3d.Architecture.Routing.Storable.Model;
using Arent3d.Revit;
using Arent3d.Utility.Serialization;
using Autodesk.Revit.DB;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
    [StorableConverterOf(typeof(CnsSettingModel))]
    internal class CnsSettingModelStorableConverter : StorableConverterBase<CnsSettingModel>
    {
        private enum SerializeField
        {
            Sequence,
            CategoryName
        }

        protected override CnsSettingModel Deserialize(Element storedElement, IDeserializerObject deserializerObject)
        {
            var deserializer = deserializerObject.Of<SerializeField>();

            int sequence = Convert.ToInt32(deserializer.GetInt(SerializeField.Sequence));
            string categoryName = deserializer.GetString(SerializeField.CategoryName)?.ToString() ?? string.Empty;

            return new CnsSettingModel(sequence, categoryName);
        }

        protected override ISerializerObject Serialize(Element storedElement, CnsSettingModel customTypeValue)
        {
            var serializerObject = new SerializerObject<SerializeField>();

            serializerObject.Add(SerializeField.Sequence, customTypeValue.Sequence);
            serializerObject.AddNonNull(SerializeField.CategoryName, customTypeValue.CategoryName);
            
            return serializerObject;
        }
    }
}
