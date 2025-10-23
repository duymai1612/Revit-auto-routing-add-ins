using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;
using System ;
using System.Collections.Generic ;
using System.Collections.ObjectModel ;
using System.Linq ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Storable.Model ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "A52AB9F8-1EB5-4BEF-9B7E-DC8CA228C12D" )]
  [StorableVisibility( AppInfo.VendorId )]
  public sealed class CnsSettingStorable : StorableBase, IEquatable<CnsSettingStorable>
  {
    public const string StorableName = "Cns Setting" ;
    private const string CnsSettingField = "CnsSetting" ;
    public ObservableCollection<CnsSettingModel> CnsSettingData { get ; set ; }

    public enum UpdateItemType
    {
      None, // デフォルト：工事項目を設定しない
      Conduit, //配線に工事項目を設定する
      Connector,
      Rack,
      All
    }

    public int SelectedIndex { get ; set ; }
    public UpdateItemType ElementType { get ; set ; }

    /// <summary>
    /// for loading from storage.
    /// </summary>
    /// <param name="owner">Owner element.</param>
    private CnsSettingStorable( DataStorage owner ) : base( owner, false )
    {
      CnsSettingData = new ObservableCollection<CnsSettingModel>() ;
      SelectedIndex = 0 ;
      ElementType = UpdateItemType.None ;
    }

    /// <summary>
    /// Called by RouteCache.
    /// </summary>
    /// <param name="document"></param>
    public CnsSettingStorable( Document document ) : base( document, false )
    {
      CnsSettingData = new ObservableCollection<CnsSettingModel>() ;
      SelectedIndex = 0 ;
      ElementType = UpdateItemType.None ;
    }

    public override string Name => StorableName ;

    protected override void LoadAllFields( FieldReader reader )
    {
      var dataSaved = reader.GetArray<CnsSettingModel>( CnsSettingField ) ;
      CnsSettingData = new ObservableCollection<CnsSettingModel>( dataSaved ) ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      writer.SetArray( CnsSettingField, CnsSettingData ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetArray<CnsSettingModel>( CnsSettingField ) ;
    }


    public bool Equals( CnsSettingStorable? other )
    {
      if ( other == null ) return false ;
      return CnsSettingData.SequenceEqual( other.CnsSettingData, new CnsSettingStorableComparer() ) ;
    }
  }

  public class CnsSettingStorableComparer : IEqualityComparer<CnsSettingModel>
  {
    public bool Equals( CnsSettingModel x, CnsSettingModel y )
    {
      return x.Equals( y ) ;
    }

    public int GetHashCode( CnsSettingModel obj )
    {
      return obj.GetHashCode() ;
    }
  }
}
