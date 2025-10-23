using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;
using Autodesk.Revit.UI.Selection ;
using MathLib ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class PointOnRoutePicker
  {
    private static AddInType? AddInType { get ; set ; }

    public class PickInfo
    {
      public Element Element { get ; }
      public XYZ Position { get ; }
      public XYZ RouteDirection { get ; }

      public Route Route => SubRoute.Route ;
      public SubRoute SubRoute { get ; }
      public Connector ReferenceConnector { get ; }

      public double Radius => GetRadius( ReferenceConnector ) ;

      public PickInfo( SubRoute subRoute, Element element, XYZ pos, XYZ dir )
      {
        SubRoute = subRoute ;
        Element = element ;
        Position = pos ;
        RouteDirection = dir ;

        ReferenceConnector = Route.GetReferenceConnector() ;
      }

      private static double GetRadius( Connector connector )
      {
        return connector.Shape switch
        {
          ConnectorProfileType.Round => connector.Radius,
          ConnectorProfileType.Oval => connector.Radius,
          ConnectorProfileType.Rectangular => 0.5 * new Vector2d( connector.Width, connector.Height ).magnitude,
          _ => 0,
        } ;
      }
    }

    public static IEnumerable<Route> PickedRoutesFromSelections( UIDocument uiDocument )
    {
      var document = uiDocument.Document ;
      var dic = RouteCache.Get( document ) ;

      var routes = new HashSet<Route>() ;

      foreach ( var elmId in uiDocument.Selection.GetElementIds() ) {
        var elm = document.GetElement( elmId ) ;
        if ( elm?.GetRouteName() is not { } routeName ) continue ;
        if ( false == dic.TryGetValue( routeName, out var route ) ) continue ;

        if ( routes.Add( route ) ) {
          yield return route ;
        }
      }
    }

    public static PickInfo PickRoute( UIDocument uiDocument, bool mepCurveOnly, string message, AddInType addInType, Predicate<Element>? elementFilter = null )
    {
      var document = uiDocument.Document ;

      var dic = RouteCache.Get( document ) ;
      AddInType = addInType ;
      var filter = new RouteFilter( dic, mepCurveOnly, elementFilter ) ;


      while ( true ) {
        var pickedObject = uiDocument.Selection.PickObject( ObjectType.PointOnElement, filter, message ) ;

        var elm = document.GetElement( pickedObject.ElementId ) ;
        if ( elm?.GetRouteName() is not { } routeName ) continue ;
        if ( false == dic.TryGetValue( routeName, out var route ) ) continue ;

        var subRoute = route.GetSubRoute( elm.GetSubRouteIndex() ?? -1 ) ;
        if ( null == subRoute ) continue ;

        var (pos, dir) = GetPositionAndDirection( elm, pickedObject.GlobalPoint ) ;
        if ( null == pos ) continue ;

        return new PickInfo( subRoute, elm, pos, dir! ) ;
      }
    }

    private static (XYZ? Position, XYZ? Direction) GetPositionAndDirection( Element elm, XYZ position )
    {
      return elm switch
      {
        MEPCurve curve => GetNearestPointAndDirection( curve, position ),
        FamilyInstance fi => ToPositionAndDirection( fi.GetTotalTransform() ),
        _ => ( null, null ),
      } ;
    }

    private static (XYZ? Position, XYZ? Direction) GetNearestPointAndDirection( MEPCurve curve, XYZ position )
    {
      var from = curve.GetRoutingConnectors( true ).FirstOrDefault() ;
      if ( null == from ) return ( null, null ) ;
      var to = curve.GetRoutingConnectors( false ).FirstOrDefault() ;
      if ( null == to ) return ( null, null ) ;

      var o = from.Origin.To3dRaw() ;
      var dir = to.Origin.To3dRaw() - o ;
      var tole = curve.Document.Application.VertexTolerance ;
      if ( dir.sqrMagnitude < tole * tole ) return ( null, null ) ;

      var line = new MathLib.Line( o, dir ) ;
      var dist = line.DistanceTo( position.To3dRaw(), 0 ) ;
      return ( Position: dist.PointOnSelf.ToXYZRaw(), Direction: dir.normalized.ToXYZRaw() ) ;
    }

    private static (XYZ Position, XYZ Direction) ToPositionAndDirection( Transform transform )
    {
      return ( transform.Origin, transform.BasisX ) ;
    }


    private class RouteFilter : ISelectionFilter
    {
      private readonly IReadOnlyDictionary<string, Route> _allRoutes ;
      private readonly bool _mepCurveOnly ;
      private readonly Predicate<Element>? _predicate ;

      public RouteFilter( IReadOnlyDictionary<string, Route> allRoutes, bool mepCurveOnly, Predicate<Element>? predicate )
      {
        _allRoutes = allRoutes ;
        _mepCurveOnly = mepCurveOnly ;
        _predicate = predicate ;
      }

      public bool AllowElement( Element elem )
      {
        if ( _mepCurveOnly && elem is not MEPCurve ) return false ;

        if ( false == elem.GetConnectors().Any( IsPickTargetConnector ) ) return false ;

        var routeName = elem.GetRouteName() ;
        if ( null == routeName ) return false ;
        if ( false == _allRoutes.ContainsKey( routeName ) ) return false ;

        return ( null == _predicate ) || _predicate( elem ) ;
      }

      public bool AllowReference( Reference reference, XYZ position )
      {
        return true ;
      }

      private static bool IsPickTargetConnector( Connector connector )
      {
        if ( AddInType == Routing.AddInType.Mechanical ) {
          return connector.IsAnyEnd() && connector.Domain switch
          {
            Domain.DomainPiping => true,
            Domain.DomainHvac => true,
            Domain.DomainElectrical => false,
            Domain.DomainCableTrayConduit => false,
            _ => false
          } ;
        }

        return connector.IsAnyEnd() && connector.Domain switch
        {
          Domain.DomainPiping => false,
          Domain.DomainHvac => false,
          Domain.DomainElectrical => true,
          Domain.DomainCableTrayConduit => true,
          _ => false
        } ;
      }
    }
  }
}