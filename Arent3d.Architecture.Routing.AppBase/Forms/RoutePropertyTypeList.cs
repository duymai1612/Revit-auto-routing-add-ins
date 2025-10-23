using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.Extensions ;
using Arent3d.Architecture.Routing.Storable ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;
using Autodesk.Revit.DB.Mechanical ;
using Autodesk.Revit.DB.Plumbing ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public class RoutePropertyTypeList
  {
    //For experimental state
    private static readonly bool UseExperimentalFeatures = true ;

    public IList<MEPSystemType>? SystemTypes { get ; }
    public IList<Opening>? Shafts { get ; }
    public IList<MEPCurveType> CurveTypes { get ; }
    public IList<string>? StandardTypes { get ; }

    public bool HasDifferentLevel { get ; }
    public (double, double) FromHeightRangeAsFloorLevel { get ; private set ; }
    public (double, double) FromHeightRangeAsCeilingLevel { get ; private set ; }
    public (double, double) ToHeightRangeAsFloorLevel { get ; private set ; }
    public (double, double) ToHeightRangeAsCeilingLevel { get ; private set ; }
    public double FromDefaultHeightAsFloorLevel { get ; private set ; }
    public double FromDefaultHeightAsCeilingLevel { get ; private set ; }
    public double ToDefaultHeightAsFloorLevel { get ; private set ; }
    public double ToDefaultHeightAsCeilingLevel { get ; private set ; }

    private void SetFromLevelSetting( HeightSettingStorable settings, ElementId levelId )
    {
      ( FromHeightRangeAsFloorLevel, FromDefaultHeightAsFloorLevel, FromHeightRangeAsCeilingLevel, FromDefaultHeightAsCeilingLevel ) = CalculateHeightRanges( settings, levelId ) ;

    }
    private void SetToLevelSetting( HeightSettingStorable settings, ElementId levelId )
    {
      ( ToHeightRangeAsFloorLevel, ToDefaultHeightAsFloorLevel, ToHeightRangeAsCeilingLevel, ToDefaultHeightAsCeilingLevel ) = CalculateHeightRanges( settings, levelId ) ;
    }

    private static ((double FloorMin, double FloorMax), double FloorDefault, (double CeilingMin, double CeilingMax), double CeilingDefault) CalculateHeightRanges( HeightSettingStorable settings, ElementId levelId )
    {
      var level = settings[ levelId ] ;
      var floorRange = ( Min: level.Underfloor.MillimetersToRevitUnits(), Max: 0d ) ;
      var ceilingRange = ( Min: 0d, Max: settings.GetDistanceToNextLevel( levelId ).MillimetersToRevitUnits() ) ;
      var floorDefault = floorRange.Min ;
      var ceilingDefault = level.HeightOfLevel.MillimetersToRevitUnits() ;
      return ( floorRange, floorDefault, ceilingRange, ceilingDefault ) ;
    }

    internal RoutePropertyTypeList( IReadOnlyCollection<SubRoute> subRoutes )
    {
      if ( 0 == subRoutes.Count ) throw new ArgumentException() ;

      var firstSubRoute = subRoutes.First() ;
      var document = firstSubRoute.Route.Document ;

      var heightSettingStorable = document.GetHeightSettingStorable() ;
      var fromLevelId = GetLevelId( document, firstSubRoute.FromEndPoints ) ;
      var toLevelId = GetLevelId( document, firstSubRoute.ToEndPoints ) ;
      HasDifferentLevel = ( fromLevelId != toLevelId ) ;
      SetFromLevelSetting( heightSettingStorable, fromLevelId != ElementId.InvalidElementId ? fromLevelId : toLevelId ) ;
      SetToLevelSetting( heightSettingStorable, toLevelId != ElementId.InvalidElementId ? toLevelId : fromLevelId ) ;

      var systemClassification = firstSubRoute.Route.GetSystemClassificationInfo() ;
      if ( systemClassification.HasSystemType() ) {
        SystemTypes = document.GetSystemTypes( systemClassification ).OrderBy( s => s.Name ).ToList() ;
        Shafts = null ;
      }
      else {
        SystemTypes = null ;
        Shafts = document.GetAllElements<Opening>().ToList() ;
      }

      CurveTypes = GetCompatibleCurveTypes( document, firstSubRoute.GetMEPCurveType().GetType() ) ;

      static ElementId GetLevelId( Document document, IEnumerable<IEndPoint> endPoints )
      {
        return endPoints.Select( ep => ep.GetLevelId( document ) ).FirstOrDefault( levelId => ElementId.InvalidElementId != levelId ) ?? ElementId.InvalidElementId ;
      }
    }

    public RoutePropertyTypeList( Document document, AddInType addInType, ElementId fromLevelId, ElementId toLevelId )
    {
      HasDifferentLevel = ( fromLevelId != toLevelId ) ;
      var heightSettingStorable = document.GetHeightSettingStorable() ;
      SetFromLevelSetting( heightSettingStorable, fromLevelId ) ;
      SetToLevelSetting( heightSettingStorable, toLevelId ) ;

      ( SystemTypes, CurveTypes, StandardTypes, Shafts ) = addInType switch
      {
        AddInType.Electrical => GetElectricalTypeLists( document ),
        AddInType.Mechanical => GetMechanicalTypeLists( document ),
        _ => throw new ArgumentOutOfRangeException( nameof( addInType ), addInType, null )
      } ;
    }

    private static (IList<MEPSystemType>? SystemTypes, IList<MEPCurveType> CurveTypes, IList<string>? StandardTypes, IList<Opening>? Shafts) GetMechanicalTypeLists( Document document )
    {
      var systemTypes = document.GetAllElements<MEPSystemType>().Where( type => type is MechanicalSystemType or PipingSystemType ).OrderBy( s => s.Name ).ToList() ;
      var curveTypes = document.GetAllElements<MEPCurveType>().Where( type => type is DuctType or PipeType ).OrderBy( s => s.Name ).ToList() ;
      return ( systemTypes, curveTypes, null, null ) ;
    }

    private static (IList<MEPSystemType>? SystemTypes, IList<MEPCurveType> CurveTypes, IList<string>? StandardTypes, IList<Opening>? Shafts) GetElectricalTypeLists( Document document )
    {
      var curveTypes = document.GetAllElements<ConduitType>().OrderBy( c => c.Name ).OfType<MEPCurveType>().ToList() ;
      var standardTypes = document.GetStandardTypes().ToList() ;
      var shafts = document.GetAllElements<Opening>().ToList() ;
      return ( null, curveTypes, standardTypes, shafts ) ;
    }

    public RoutePropertyTypeList( Document document, MEPSystemClassificationInfo classificationInfo, ElementId fromLevelId, ElementId toLevelId )
    {
      HasDifferentLevel = ( fromLevelId != toLevelId ) ;
      var heightSettingStorable = document.GetHeightSettingStorable() ;
      SetFromLevelSetting( heightSettingStorable, fromLevelId ) ;
      SetToLevelSetting( heightSettingStorable, toLevelId ) ;

      if ( classificationInfo.HasSystemType() ) {
        SystemTypes = document.GetSystemTypes( classificationInfo ).OrderBy( s => s.Name ).ToList() ;
        CurveTypes = GetCompatibleCurveTypes( document, classificationInfo.GetCurveTypeClass() ) ;
      }
      else {
        CurveTypes = document.GetAllElements<ConduitType>().OrderBy( c => c.Name ).OfType<MEPCurveType>().ToList() ;
        StandardTypes = document.GetStandardTypes().ToList() ;
        Shafts = document.GetAllElements<Opening>().ToList() ;
      }
    }

    private static IList<MEPCurveType> GetCompatibleCurveTypes( Document document, Type? mepCurveTypeClass )
    {
      var curveTypes = document.GetCurveTypes( mepCurveTypeClass ) ;
      if ( UseExperimentalFeatures ) {
        curveTypes = curveTypes.Where( c => c.Shape == ConnectorProfileType.Round ) ;
      }

      return curveTypes.OrderBy( s => s.Name ).ToList() ;
    }
  }
  
}