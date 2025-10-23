using System.Collections.Generic ;
using System.Linq ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "2abee280-4a54-4256-945f-ca5fc4f57ab3" )]
  [StorableVisibility( AppInfo.VendorId )]
  public class CsvStorable : StorableBase
  {
    public const string StorableName = "Csv Model" ;
    private const string WiresAndCablesModelField = "WiresAndCablesModel" ;
    private const string ConduitsModelField = "ConduitsModel" ;
    private const string HiroiSetMasterNormalModelField = "HiroiSetMasterNormalModelModel" ;
    private const string HiroiSetMasterEcoModelField = "HiroiSetMasterEcoModel" ;
    private const string HiroiSetCdMasterNormalModelField = "HiroiSetCdMasterNormalModelModel" ;
    private const string HiroiSetCdMasterEcoModelField = "HiroiSetCdMasterEcoModel" ;
    private const string HiroiMasterModelField = "HiroiMasterEcoModel" ;

    public List<WiresAndCablesModel> WiresAndCablesModelData { get ; set ; }
    public List<ConduitsModel> ConduitsModelData { get ; set ; }
    public List<HiroiSetMasterModel> HiroiSetMasterNormalModelData { get ; set ; }
    public List<HiroiSetMasterModel> HiroiSetMasterEcoModelData { get ; set ; }
    public List<HiroiSetCdMasterModel> HiroiSetCdMasterNormalModelData { get ; set ; }
    public List<HiroiSetCdMasterModel> HiroiSetCdMasterEcoModelData { get ; set ; }
    public List<HiroiMasterModel> HiroiMasterModelData { get ; set ; }

    public CsvStorable( DataStorage owner ) : base( owner, false )
    {
      WiresAndCablesModelData = new List<WiresAndCablesModel>() ;
      ConduitsModelData = new List<ConduitsModel>() ;
      HiroiSetMasterNormalModelData = new List<HiroiSetMasterModel>() ;
      HiroiSetMasterEcoModelData = new List<HiroiSetMasterModel>() ;
      HiroiSetCdMasterNormalModelData = new List<HiroiSetCdMasterModel>() ;
      HiroiSetCdMasterEcoModelData = new List<HiroiSetCdMasterModel>() ;
      HiroiMasterModelData = new List<HiroiMasterModel>() ;
    }

    public CsvStorable( Document document ) : base( document, false )
    {
      WiresAndCablesModelData = new List<WiresAndCablesModel>() ;
      ConduitsModelData = new List<ConduitsModel>() ;
      HiroiSetMasterNormalModelData = new List<HiroiSetMasterModel>() ;
      HiroiSetMasterEcoModelData = new List<HiroiSetMasterModel>() ;
      HiroiSetCdMasterNormalModelData = new List<HiroiSetCdMasterModel>() ;
      HiroiSetCdMasterEcoModelData = new List<HiroiSetCdMasterModel>() ;
      HiroiMasterModelData = new List<HiroiMasterModel>() ;
    }

    protected override void LoadAllFields( FieldReader reader )
    {
      WiresAndCablesModelData = reader.GetArray<WiresAndCablesModel>( WiresAndCablesModelField ).ToList() ;
      ConduitsModelData = reader.GetArray<ConduitsModel>( ConduitsModelField ).ToList() ;
      HiroiSetMasterNormalModelData = reader.GetArray<HiroiSetMasterModel>( HiroiSetMasterNormalModelField ).ToList() ;
      HiroiSetMasterEcoModelData = reader.GetArray<HiroiSetMasterModel>( HiroiSetMasterEcoModelField ).ToList() ;
      HiroiSetCdMasterNormalModelData = reader.GetArray<HiroiSetCdMasterModel>( HiroiSetCdMasterNormalModelField ).ToList() ;
      HiroiSetCdMasterEcoModelData = reader.GetArray<HiroiSetCdMasterModel>( HiroiSetCdMasterEcoModelField ).ToList() ;
      HiroiMasterModelData = reader.GetArray<HiroiMasterModel>( HiroiMasterModelField ).ToList() ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      writer.SetArray( WiresAndCablesModelField, WiresAndCablesModelData ) ;
      writer.SetArray( ConduitsModelField, ConduitsModelData ) ;
      writer.SetArray( HiroiSetMasterNormalModelField, HiroiSetMasterNormalModelData ) ;
      writer.SetArray( HiroiSetMasterEcoModelField, HiroiSetMasterEcoModelData ) ;
      writer.SetArray( HiroiSetCdMasterNormalModelField, HiroiSetCdMasterNormalModelData ) ;
      writer.SetArray( HiroiSetCdMasterEcoModelField, HiroiSetCdMasterEcoModelData ) ;
      writer.SetArray( HiroiMasterModelField, HiroiMasterModelData ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetArray<WiresAndCablesModel>( WiresAndCablesModelField ) ;
      generator.SetArray<ConduitsModel>( ConduitsModelField ) ;
      generator.SetArray<HiroiSetMasterModel>( HiroiSetMasterNormalModelField ) ;
      generator.SetArray<HiroiSetMasterModel>( HiroiSetMasterEcoModelField ) ;
      generator.SetArray<HiroiSetCdMasterModel>( HiroiSetCdMasterNormalModelField ) ;
      generator.SetArray<HiroiSetCdMasterModel>( HiroiSetCdMasterEcoModelField ) ;
      generator.SetArray<HiroiMasterModel>( HiroiMasterModelField ) ;
    }

    public override string Name => StorableName ;
  }
}