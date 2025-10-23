using Arent3d.Architecture.Routing.Storable ;
using Arent3d.Architecture.Routing.Storable.StorableConverter ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.Exceptions ;

namespace Arent3d.Architecture.Routing.Extensions
{
  public static class DocumentExtensions
  {
    /// <summary>
    /// Get Height settings data from snoop DB. <br />
    /// If there is no data, it is returned default settings
    /// </summary>
    /// <param name="document">current document of Revit</param>
    /// <returns>Height settings data was stored in snoop DB</returns>
    public static HeightSettingStorable GetHeightSettingStorable( this Document document )
    {
      try {
        return HeightSettingStorableCache.Get( document ).FindOrCreate( HeightSettingStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new HeightSettingStorable( document ) ;
      }
    }

    /// <summary>
    /// Get Offset settings data from snoop DB.
    /// </summary>
    public static OffsetSettingStorable GetOffsetSettingStorable( this Document document )
    {
      try {
        return OffsetSettingStorableCache.Get( document ).FindOrCreate( OffsetSettingStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new OffsetSettingStorable( document ) ;
      }
    }
    
    /// <summary>
    /// Get CNS Setting data from snoop DB.
    /// </summary>
    public static CnsSettingStorable GetCnsSettingStorable( this Document document )
    {
      try {
        return CnsSettingStorableCache.Get( document ).FindOrCreate( CnsSettingStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new CnsSettingStorable( document ) ;
      }
    }

    /// <summary>
    /// Get Ceed Model data from snoop DB.
    /// </summary>
    public static CeedStorable GetCeeDStorable( this Document document )
    {
      try {
        return CeedStorableCache.Get( document ).FindOrCreate( CeedStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new CeedStorable( document ) ;
      }
    }

    /// <summary>
    /// Get csv data from snoop DB.
    /// </summary>
    public static CsvStorable GetCsvStorable( this Document document )
    {
      try {
        return CsvStorableCache.Get( document ).FindOrCreate( CsvStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new CsvStorable( document ) ;
      }
    }

    /// <summary>
    /// Get pick up data from snoop DB.
    /// </summary>
    public static PickUpStorable GetPickUpStorable( this Document document )
    {
      try {
        return PickUpStorableCache.Get( document ).FindOrCreate( PickUpStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new PickUpStorable( document ) ;
      }
    }
    
    /// <summary>
    /// Get detail symbol data from snoop DB.
    /// </summary>
    public static DetailSymbolStorable GetDetailSymbolStorable( this Document document )
    {
      try {
        return DetailSymbolStorableCache.Get( document ).FindOrCreate( DetailSymbolStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new DetailSymbolStorable( document ) ;
      }
    }
    
    /// <summary>
    /// Get rack notation data from snoop DB.
    /// </summary>
    public static RackNotationStorable GetRackNotationStorable( this Document document )
    {
      try {
        return RackNotationStorableCache.Get( document ).FindOrCreate( RackNotationStorable.StorableName ) ;
      }
      catch ( InvalidOperationException ) {
        return new RackNotationStorable( document ) ;
      }
    }
  }
}
