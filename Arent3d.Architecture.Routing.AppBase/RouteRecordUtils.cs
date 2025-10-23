using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class RouteRecordUtils
  {
    public static IEnumerable<(string RouteName, RouteSegment Segment)> ToSegmentsWithName( this IEnumerable<RouteSegment> segments, string routeName )
    {
      return segments.Select( seg => ( routeName, seg ) ) ;
    }
    public static IEnumerable<(string RouteName, RouteSegment Segment)> ToSegmentsWithName( this Route route )
    {
      return route.RouteSegments.ToSegmentsWithName( route.RouteName ) ;
    }
    public static List<(string RouteName, RouteSegment Segment)> ToSegmentsWithNameList( this Route route )
    {
      var list = new List<(string RouteName, RouteSegment Segment)>( route.RouteSegments.Count ) ;
      list.AddRange( route.RouteSegments.ToSegmentsWithName( route.RouteName ) ) ;
      return list ;
    }
    public static IEnumerable<(string RouteName, RouteSegment Segment)> ToSegmentsWithName( this IEnumerable<Route> routes )
    {
      return routes.SelectMany( ToSegmentsWithName ) ;
    }

    public static IEnumerable<(string RouteName, RouteSegment Segment)> ToSegmentsWithName( this IEnumerable<RouteRecord> routeRecords, Document document )
    {
      var endPointDictionary = new EndPointDictionaryForImport( document ) ;

      foreach ( var record in routeRecords ) {
        var fromEndPoint = endPointDictionary.GetEndPoint( record.RouteName, record.FromKey,document.ParseEndPoint( record.FromEndType, record.FromEndParams ) ) ;
        var toEndPoint = endPointDictionary.GetEndPoint( record.RouteName, record.ToKey, document.ParseEndPoint( record.ToEndType, record.ToEndParams ) ) ;
        if ( null == fromEndPoint || null == toEndPoint ) continue ;

        var classificationInfo = MEPSystemClassificationInfo.Deserialize( record.SystemClassification ) ;
        if ( null == classificationInfo ) continue ;

        var systemType = GetSystemType( document, classificationInfo, record.SystemTypeName ) ;
        if ( null == systemType ) continue ;

        var curveTypeClass = classificationInfo.GetCurveTypeClass() ;
        MEPCurveType? curveType = null ;
        if ( null != curveTypeClass ) {
          curveType = GetCurveType( document, curveTypeClass, record.CurveTypeName ) ;
          if ( null != curveType ) continue ;
        }
        
        var fromFixedHeight=FixedHeight.CreateOrNull( record.FromFixedHeightType, record.FromFixedHeightValue );
        var toFixedHeight=FixedHeight.CreateOrNull( record.ToFixedHeightType, record.ToFixedHeightValue );

        yield return ( record.RouteName, new RouteSegment( classificationInfo, systemType, curveType, fromEndPoint, toEndPoint, record.NominalDiameter, record.IsRoutingOnPipeSpace, fromFixedHeight, toFixedHeight, record.AvoidType, ToElementId( record.ShaftElementId ) ) ) ;
      }
    }

    private static ElementId ToElementId( int elementId )
    {
      if ( elementId <= 0 ) return ElementId.InvalidElementId ;
      return new ElementId( elementId ) ;
    }

    private static MEPSystemType? GetSystemType( Document document, MEPSystemClassificationInfo classificationInfo, string systemTypeName )
    {
      var allSystemTypes = document.GetAllElements<MEPSystemType>().Where( classificationInfo.IsCompatibleTo ).EnumerateAll() ;
      return allSystemTypes.FirstOrDefault( systemType => systemType.Name == systemTypeName ) ?? allSystemTypes.FirstOrDefault() ;
    }

    private static MEPCurveType? GetCurveType( Document document, Type curveTypeClass, string curveTypeName )
    {
      var allCurveTypes = document.GetAllElements<MEPCurveType>( curveTypeClass ).EnumerateAll() ;
      return allCurveTypes.FirstOrDefault( ct => ct.Name == curveTypeName ) ?? allCurveTypes.FirstOrDefault() ;
    }

    public static IEnumerable<RouteRecord> ToRouteRecords( this IEnumerable<(string RouteName, RouteSegment Segment)> segments, Document document )
    {
      var endPointDictionary = new EndPointDictionaryForExport( document ) ;

      foreach ( var (routeName, segment) in segments ) {
        var (fromKey, fromEndPoint) = endPointDictionary.GetEndPoint( segment.FromEndPoint ) ;
        var (toKey, toEndPoint) = endPointDictionary.GetEndPoint( segment.ToEndPoint ) ;

        yield return new RouteRecord
        {
          RouteName = routeName,
          FromKey = fromKey,
          FromEndType = fromEndPoint.TypeName,
          FromEndParams = fromEndPoint.ParameterString,
          ToKey = toKey,
          ToEndType = toEndPoint.TypeName,
          ToEndParams = toEndPoint.ParameterString,
          NominalDiameter = segment.PreferredNominalDiameter,
          CurveTypeName = GetCurveTypeName( segment.CurveType ),
          FromFixedHeightType = segment.FromFixedHeight?.Type.ToString() ?? string.Empty,
          FromFixedHeightValue = segment.FromFixedHeight?.Height,
          ToFixedHeightType = segment.ToFixedHeight?.Type.ToString() ?? string.Empty,
          ToFixedHeightValue = segment.ToFixedHeight?.Height,
          AvoidType = segment.AvoidType,
          SystemClassification = segment.SystemClassificationInfo.Serialize(),
          SystemTypeName = GetSystemTypeName( segment.SystemType ),
          ShaftElementId = segment.ShaftElementId.IntegerValue,
        } ;
      }
    }

    private static string GetCurveTypeName( MEPCurveType? curveType )
    {
      if ( null == curveType ) return string.Empty ;

      return curveType.Name ;
    }
    private static string GetSystemTypeName( MEPSystemType? systemType )
    {
      if ( null == systemType ) return string.Empty ;

      return systemType.Name ;
    }
  }
}