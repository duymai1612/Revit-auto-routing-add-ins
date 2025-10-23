using System.Collections.Generic ;
using System.Linq ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "27e3cb9a-bdbc-4ce5-9f72-9ce6c2ea5c13" )]
  [StorableVisibility( AppInfo.VendorId )]
  public class RackNotationStorable : StorableBase
  {
    public const string StorableName = "Rack Notation Model" ;
    private const string RackNotationModelField = "RackNotationModel" ;
    public List<RackNotationModel> RackNotationModelData { get ; set ; }

    public RackNotationStorable( DataStorage owner ) : base( owner, false )
    {
      RackNotationModelData = new List<RackNotationModel>() ;
    }

    public RackNotationStorable( Document document ) : base( document, false )
    {
      RackNotationModelData = new List<RackNotationModel>() ;
    }

    protected override void LoadAllFields( FieldReader reader )
    {
      RackNotationModelData = reader.GetArray<RackNotationModel>( RackNotationModelField ).ToList() ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      writer.SetArray( RackNotationModelField, RackNotationModelData ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetArray<PickUpModel>( RackNotationModelField ) ;
    }

    public override string Name => StorableName ;
  }
}