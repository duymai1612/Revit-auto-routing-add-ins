using System.Collections.Generic ;
using System.Linq ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "52c3c670-d3b0-4dcc-bc27-871c90bf173e" )]
  [StorableVisibility( AppInfo.VendorId )]
  public class DetailSymbolStorable : StorableBase
  {
    public const string StorableName = "Detail Symbol Model" ;
    private const string DetailSymbolModelField = "DetailSymbolModel" ;
    public List<DetailSymbolModel> DetailSymbolModelData { get ; set ; }
    
    public DetailSymbolStorable( DataStorage owner ) : base( owner, false )
    {
      DetailSymbolModelData = new List<DetailSymbolModel>() ;
    }

    public DetailSymbolStorable( Document document ) : base( document, false )
    {
      DetailSymbolModelData = new List<DetailSymbolModel>() ;
    }

    protected override void LoadAllFields( FieldReader reader )
    {
      DetailSymbolModelData = reader.GetArray<DetailSymbolModel>( DetailSymbolModelField ).ToList() ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      writer.SetArray( DetailSymbolModelField, DetailSymbolModelData ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetArray<DetailSymbolModel>( DetailSymbolModelField ) ;
    }

    public override string Name => StorableName ;
  }
}