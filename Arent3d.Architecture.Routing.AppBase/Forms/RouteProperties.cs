using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public class RouteProperties
  {
    private Document Document { get ; }
    public double VertexTolerance => Document.Application.VertexTolerance ;

    //Diameter
    public double? Diameter { get ; private set ; }

    //SystemType 
    public MEPSystemType? SystemType { get ; }
    
    // Shaft 
    public Opening? Shaft { get ; private set ; }    

    //CurveType
    public MEPCurveType? CurveType { get ; private set ; }

    //Direct
    public bool? IsRouteOnPipeSpace { get ; private set ; }

    //HeightSetting
    public bool? UseFromFixedHeight { get ; private set ; }
    public FixedHeight? FromFixedHeight { get ; private set ; }
    public bool? UseToFixedHeight { get ; private set ; }
    public FixedHeight? ToFixedHeight { get ; private set ; }

    public AvoidType? AvoidType { get ; private set ; }

    public string? StandardType { get ; }

    internal RouteProperties( IReadOnlyCollection<SubRoute> subRoutes )
    {
      if ( 0 == subRoutes.Count ) throw new ArgumentException() ;

      var firstSubRoute = subRoutes.First() ;
      var document = firstSubRoute.Route.Document ;

      //Set properties
      CurveType = firstSubRoute.GetMEPCurveType() ;

      //Diameter Info
      Document = document ;
      Diameter = firstSubRoute.GetDiameter() ;

      //System Type Info(PipingSystemType in lookup)
      var systemClassification = firstSubRoute.Route.GetSystemClassificationInfo() ;
      if ( systemClassification.HasSystemType() ) {
        SystemType = firstSubRoute.Route.GetMEPSystemType() ;
      }
      else {
        SystemType = null ;
      }

      //Direct Info
      IsRouteOnPipeSpace = firstSubRoute.IsRoutingOnPipeSpace ;

      //Height Info
      UseFromFixedHeight = ( firstSubRoute.FromFixedHeight != null ) ;
      FromFixedHeight = firstSubRoute.FromFixedHeight ;
      UseToFixedHeight = ( firstSubRoute.ToFixedHeight != null ) ;
      ToFixedHeight = firstSubRoute.ToFixedHeight ;

      //AvoidType Info
      AvoidType = firstSubRoute.AvoidType ;

      if ( 1 < subRoutes.Count ) {
        SetIndeterminateValues( firstSubRoute.Route ) ;
      }

      Shaft = document.GetElementById<Opening>( firstSubRoute.ShaftElementId ) ;
    }

    public RouteProperties( Document document, RoutePropertyTypeList spec )
    {
      Document = document ;

      SystemType = spec.SystemTypes?.FirstOrDefault() ;
      Shaft = spec.Shafts?.FirstOrDefault() ;
      CurveType = spec.CurveTypes.FirstOrDefault() ;
      Diameter = null ;

      IsRouteOnPipeSpace = false ;
      UseFromFixedHeight = true ;
      UseToFixedHeight = false ;
      FromFixedHeight = null ;
      ToFixedHeight = null ;
      AvoidType = Routing.AvoidType.Whichever ;
    }

    public RouteProperties( Document document, MEPSystemClassificationInfo classificationInfo, MEPSystemType? systemType, MEPCurveType? curveType, string? standardType, double? diameter )
    {
      Document = document ;
      
      Diameter = diameter ;

      if ( classificationInfo.HasSystemType() ) {
        // Mechanical
        SystemType = systemType ;
        CurveType = curveType ;

        IsRouteOnPipeSpace = false ;
        UseFromFixedHeight = true ;
        UseToFixedHeight = false ;
        FromFixedHeight = null ;
        ToFixedHeight = null ;
        AvoidType = Routing.AvoidType.Whichever ;
      }
      else {
        // Electrical
        CurveType = curveType ;
        StandardType = standardType ;

        IsRouteOnPipeSpace = false ;
        UseFromFixedHeight = true ;
        UseToFixedHeight = false ;
        FromFixedHeight = null ;
        ToFixedHeight = null ;
        AvoidType = Routing.AvoidType.Whichever ;
      }
    }

    public RouteProperties( Route route, MEPSystemType? systemType, MEPCurveType? curveType, double? diameter, bool? isRouteOnPipeSpace, bool? useFromFixedHeight, FixedHeight? fromFixedHeight, bool? useToFixedHeight, FixedHeight? toFixedHeight, AvoidType? avoidType, Opening? shaft )
      : this( route.Document, systemType, curveType, diameter, isRouteOnPipeSpace, useFromFixedHeight, fromFixedHeight, useToFixedHeight, toFixedHeight, avoidType, shaft ){}
    public RouteProperties( Document document, MEPSystemType? systemType, MEPCurveType? curveType, double? diameter, bool? isRouteOnPipeSpace, bool? useFromFixedHeight, FixedHeight? fromFixedHeight, bool? useToFixedHeight, FixedHeight? toFixedHeight, AvoidType? avoidType, Opening? shaft )
    {
      Document = document ;

      Diameter = diameter ;

      CurveType = curveType ;

      SystemType = systemType ;

      IsRouteOnPipeSpace = isRouteOnPipeSpace ;
      UseFromFixedHeight = useFromFixedHeight ;
      if ( true == UseFromFixedHeight ) {
        FromFixedHeight = fromFixedHeight ;
      }
      UseToFixedHeight = useToFixedHeight ;
      if ( true == UseToFixedHeight ) {
        ToFixedHeight = toFixedHeight ;
      }
      AvoidType = avoidType ;
      Shaft = shaft ;
    }

    private void SetIndeterminateValues( Route route )
    {
      // if Diameter is multi selected, set null
      if ( IsDiameterMultiSelected( route ) ) {
        Diameter = null ;
      }

      // if CurveType is multi selected, set null
      if ( IsCurveTypeMultiSelected( route ) ) {
        CurveType = null ;
      }

      IsRouteOnPipeSpace = route.UniqueIsRoutingOnPipeSpace ;

      if ( IsUseFromFixedHeightMultiSelected( route ) ) {
        UseFromFixedHeight = null ;
        FromFixedHeight = null ;
      }
      if ( IsUseToFixedHeightMultiSelected( route ) ) {
        UseToFixedHeight = null ;
        ToFixedHeight = null ;
      }

      if ( IsAvoidTypeMultiSelected( route ) ) {
        AvoidType = null ;
      }

      if ( IsShaftMultiSelected( route ) ) {
        Shaft = null ;
      }
    }

    /// <summary>
    /// Get CurveType's multi selected state
    /// </summary>
    /// <param name="route"></param>
    /// <returns></returns>
    private static bool IsCurveTypeMultiSelected( Route route )
    {
      return ( null == route.UniqueCurveType ) ;
    }

    /// <summary>
    /// Get Diameter's multi selected state
    /// </summary>
    /// <param name="route"></param>
    /// <returns></returns>
    private static bool IsDiameterMultiSelected( Route route )
    {
      return ( null == route.UniqueDiameter ) ;
    }

    private static bool IsUseFromFixedHeightMultiSelected( Route route )
    {
      if ( route.SubRoutes.Any( subRoute => null != subRoute.FromFixedHeight ) ) return false ;

      return ( null == route.UniqueFromFixedHeight ) ;
    }
    private static bool IsUseToFixedHeightMultiSelected( Route route )
    {
      if ( route.SubRoutes.Any( subRoute => null != subRoute.ToFixedHeight ) ) return false ;

      return ( null == route.UniqueToFixedHeight ) ;
    }

    /// <summary>
    /// Get AvoidType's multi selected state
    /// </summary>
    /// <param name="route"></param>
    /// <returns></returns>
    private static bool IsAvoidTypeMultiSelected( Route route )
    {
      return ( null == route.UniqueAvoidType ) ;
    }

    /// <summary>
    /// Get Shaft's multi selected state
    /// </summary>
    /// <param name="route"></param>
    /// <returns></returns>
    private static bool IsShaftMultiSelected( Route route )
    {
      return ( null == route.UniqueShaftElementId ) ;
    }
  }
}