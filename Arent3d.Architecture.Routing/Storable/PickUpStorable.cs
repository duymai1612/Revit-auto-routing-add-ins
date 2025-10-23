using System.Collections.Generic ;
using System.Linq ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "14e49565-23e8-48f2-9873-653f285b6901" )]
  [StorableVisibility( AppInfo.VendorId )]
  public class PickUpStorable : StorableBase
  {
    public const string StorableName = "Pick-Up Model" ;
    private const string AllPickUpModelField = "AllPickUpModel" ;
    public List<PickUpModel> AllPickUpModelData { get ; set ; }

    public PickUpStorable( DataStorage owner ) : base( owner, false )
    {
      AllPickUpModelData = new List<PickUpModel>() ;
    }

    public PickUpStorable( Document document ) : base( document, false )
    {
      AllPickUpModelData = new List<PickUpModel>() ;
    }

    protected override void LoadAllFields( FieldReader reader )
    {
      AllPickUpModelData = reader.GetArray<PickUpModel>( AllPickUpModelField ).ToList() ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      writer.SetArray( AllPickUpModelField, AllPickUpModelData ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetArray<PickUpModel>( AllPickUpModelField ) ;
    }

    public override string Name => StorableName ;
  }
}