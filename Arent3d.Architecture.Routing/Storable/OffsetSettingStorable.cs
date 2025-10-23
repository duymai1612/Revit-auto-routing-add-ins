using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;
using System ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Storable.Model ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "7941878b-c02e-4694-856c-0dd67655a76e" )]
  [StorableVisibility( AppInfo.VendorId )]
  public sealed class OffsetSettingStorable : StorableBase
  {
    public const string StorableName = "Offset Setting" ;
    private const string OffsetSettingField = "OffsetSetting" ;

    public OffsetSettingModel OffsetSettingsData { get ; private set ; }


    /// <summary>
    /// For loading from storage.
    /// </summary>
    private OffsetSettingStorable( DataStorage owner ) : base( owner, false )
    {
      OffsetSettingsData = new OffsetSettingModel() ;
    }

    /// <summary>
    /// Called by RouteCache.
    /// </summary>
    public OffsetSettingStorable( Document document ) : base( document, false )
    {
      OffsetSettingsData = new OffsetSettingModel() ;
    }

    public override string Name => StorableName ;


    protected override void LoadAllFields( FieldReader reader )
    {
      var dataSaved = reader.GetSingle<OffsetSettingModel>( OffsetSettingField ) ;

      OffsetSettingsData = dataSaved ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      writer.SetSingle( OffsetSettingField, OffsetSettingsData ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetSingle<OffsetSettingModel>( OffsetSettingField ) ;
    }
  }
}