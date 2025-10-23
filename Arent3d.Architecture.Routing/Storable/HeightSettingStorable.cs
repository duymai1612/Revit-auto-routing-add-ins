using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.ExtensibleStorage ;
using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Runtime.InteropServices ;
using Arent3d.Architecture.Routing.Extensions ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Utility ;

namespace Arent3d.Architecture.Routing.Storable
{
  [Guid( "83A448F4-E120-44E0-A220-F2D3F11B6A09" )]
  [StorableVisibility( AppInfo.VendorId )]
  public sealed class HeightSettingStorable : StorableBase, IEquatable<HeightSettingStorable>
  {
    public const string StorableName = "Height Setting" ;
    private const double DefaultMaxLevelDistance = 100000 ; // max level distance
    
    private const string HeightSettingField = "HeightSetting" ;

    public Dictionary<int, HeightSettingModel> HeightSettingsData { get ; private set ; }
    public IReadOnlyList<Level> Levels { get ; }

    /// <summary>
    /// Get Height settings data by Level object
    /// </summary>
    /// <param name="level"></param>
    public HeightSettingModel this[ Level level ] => HeightSettingsData.GetOrDefault( level.GetValidId().IntegerValue, () => new HeightSettingModel( level ) ) ;

    /// <summary>
    /// Get Height settings data by level Id.
    /// </summary>
    /// <param name="levelId"></param>
    public HeightSettingModel this[ int levelId ]
    {
      get
      {
        var levelIndex = Levels.FindIndex( x => x.GetValidId().IntegerValue == levelId ) ;
        if ( levelIndex < 0 ) throw new KeyNotFoundException() ;
        return this[ Levels[ levelIndex ] ] ;
      }
    }

    /// <summary>
    /// Get Height settings data by level Id.
    /// </summary>
    /// <param name="levelId"></param>
    public HeightSettingModel this[ ElementId levelId ] => this[ levelId.IntegerValue ] ;


    /// <summary>
    /// for loading from storage.
    /// </summary>
    /// <param name="owner">Owner element.</param>
    private HeightSettingStorable( DataStorage owner ) : base( owner, false )
    {
      Levels = GetAllLevels( owner.Document ) ;
      HeightSettingsData = new Dictionary<int, HeightSettingModel>() ;
    }

    /// <summary>
    /// Called by RouteCache.
    /// </summary>
    /// <param name="document"></param>
    public HeightSettingStorable( Document document ) : base( document, false )
    {
      Levels = GetAllLevels( document ) ;
      HeightSettingsData = Levels.ToDictionary( x => x.Id.IntegerValue, x => new HeightSettingModel( x ) ) ;
    }

    private static IReadOnlyList<Level> GetAllLevels( Document document )
    {
      var levels = document.GetAllElements<Level>().ToList() ;
      levels.Sort( ( a, b ) => a.Elevation.CompareTo( b.Elevation ) ) ;
      return levels ;
    }

    public override string Name => StorableName ;

    public double GetAbsoluteHeight( ElementId levelId, FixedHeightType fixedHeightType, double fixedHeightHeight )
    {
      return this[ levelId ].Elevation.MillimetersToRevitUnits() + fixedHeightHeight ;
    }

    public double GetDistanceToNextLevel( ElementId levelId )
    {
      var index = Levels.FindIndex( level => level.GetValidId() == levelId ) ;
      if ( index < 0 || Levels.Count - 1 <= index ) return DefaultMaxLevelDistance ;

      return this[ Levels[ index + 1 ] ].Elevation - this[ Levels[ index ] ].Elevation ;
    }

    protected override void LoadAllFields( FieldReader reader )
    {
      var dataSaved = reader.GetArray<HeightSettingModel>( HeightSettingField ).ToDictionary( x => x.LevelId ) ;

      HeightSettingsData = Levels.ToDictionary( x => x.Id.IntegerValue, x => dataSaved.GetOrDefault( x.Id.IntegerValue, () => new HeightSettingModel( x ) ) ) ;
    }

    protected override void SaveAllFields( FieldWriter writer )
    {
      HeightSettingsData = Levels.ToDictionary( x => x.Id.IntegerValue, x => HeightSettingsData.GetOrDefault( x.Id.IntegerValue, () => new HeightSettingModel( x ) ) ) ;

      writer.SetArray( HeightSettingField, HeightSettingsData.Values.ToList() ) ;
    }

    protected override void SetupAllFields( FieldGenerator generator )
    {
      generator.SetArray<HeightSettingModel>( HeightSettingField ) ;
    }

    public bool Equals( HeightSettingStorable? other )
    {
      if ( other == null ) return false ;
      return HeightSettingsData.Values.OrderBy( x => x.LevelId ).SequenceEqual( other.HeightSettingsData.Values.OrderBy( x => x.LevelId ), new HeightSettingStorableComparer() ) ;
    }
  }

  public class HeightSettingStorableComparer : IEqualityComparer<HeightSettingModel>
  {
    public bool Equals( HeightSettingModel x, HeightSettingModel y )
    {
      return x.Equals( y ) ;
    }

    public int GetHashCode( HeightSettingModel obj )
    {
      return obj.GetHashCode() ;
    }
  }

}