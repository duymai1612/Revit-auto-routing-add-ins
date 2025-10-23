using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.AppBase.Forms ;
using Arent3d.Architecture.Routing.AppBase.UI ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;
using Autodesk.Revit.UI.Selection ;
using Arent3d.Revit.UI ;
using Arent3d.Utility ;
using Autodesk.Revit.DB.Architecture ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class ConnectorPicker
  {
    public interface IPickResult
    {
      IEnumerable<ElementId> GetAllRelatedElements() ;
      ElementId GetLevelId() ;
      SubRoute? SubRoute { get ; }
      EndPointKey? EndPointOverSubRoute { get ; }
      Element PickedElement { get ; }
      Connector? PickedConnector { get ; }
      XYZ GetOrigin() ;
      XYZ? GetMEPCurveDirection( bool isFrom ) ;
      Element GetOriginElement() ;
      bool IsCompatibleTo( Connector connector ) ;
      bool IsCompatibleTo( Element element ) ;
    }

    public static IPickResult GetConnector( UIDocument uiDocument, RoutingExecutor routingExecutor, bool pickingFromSide, string message, IPickResult? compatiblePickResult, AddInType addInType )
    {
      var document = uiDocument.Document ;

      var filter = ( null == compatiblePickResult ) ? FamilyInstanceWithConnectorFilter.GetInstance( addInType ) : new FamilyInstanceCompatibleToTargetConnectorFilter( compatiblePickResult, addInType ) ;

      while ( true ) {
        var pickedObject = uiDocument.Selection.PickObject( ObjectType.Element, filter, message ) ;

        var element = document.GetElement( pickedObject.ElementId ) ;
        if ( null == element ) continue ;

        if ( PassPointPickResult.Create( element ) is { } ppResult ) return ppResult ;
        if ( AddInType.Electrical == addInType ) {
          if ( SubRoutePickResult.Create( routingExecutor, element, pickedObject.GlobalPoint ) is {} srResult ) {
            if ( PickEndPointOverSubRoute( uiDocument, srResult, pickingFromSide ) is not { } endPoint ) continue ;
            return srResult.ApplyEndPointOverSubRoute( endPoint.Key ) ;
          }
        }
        else {
          if ( SubRoutePickResult.Create( element, pickedObject.GlobalPoint ) is { } srResult ) return srResult ;
        }

        var conn = compatiblePickResult?.SubRoute?.GetReferenceConnector() ?? compatiblePickResult?.PickedConnector ;

        var (result, connector) = FindConnector( uiDocument, element, message, conn, addInType ) ;
        if ( false == result ) continue ;

        if ( null != connector ) {
          return new ConnectorPickResult( element, connector ) ;
        }

        return new OriginPickResult( element, addInType ) ;
      }
    }

    private static IEndPoint? PickEndPointOverSubRoute( UIDocument uiDocument, SubRoutePickResult pickResult, bool pickingFromSide )
    {
      var endPoints = new Dictionary<EndPointKey, IEndPoint>() ;
      foreach ( var endPoint in GetNearestEndPoints( pickResult.GetOriginElement(), pickingFromSide ) ) {
        var key = endPoint.Key ;
        if ( endPoints.ContainsKey( key ) ) continue ;
        endPoints.Add( key, endPoint ) ;
      }

      if ( 0 == endPoints.Count ) return null ;
      if ( 1 == endPoints.Count ) return endPoints.Values.FirstOrDefault() ;

      using var displayEndPoints = new EndPointPicker( uiDocument, pickResult.GetAllRelatedElements(), endPoints.Values ) ;
      return displayEndPoints.Pick() ;
    }

    private static IEnumerable<IEndPoint> GetNearestEndPoints( Element originElement, bool pickingFromSide )
    {
      if ( originElement is not MEPCurve mepCurve ) return Array.Empty<IEndPoint>() ;

      var document = originElement.Document ;
      return mepCurve.GetSubRouteGroup().Select( subRoute => GetEndPoint( document, subRoute, pickingFromSide ) ).NonNull() ;

      static IEndPoint? GetEndPoint( Document document, SubRoute subRoute, bool pickingFromSide )
      {
        while ( true ) {  // tail recursion
          if ( pickingFromSide ) {
            var endPoint = subRoute.Segments.FirstOrDefault()?.FromEndPoint.ToEndPointOverSubRoute( document ) ;
            if ( endPoint is not PassPointEndPoint ) return endPoint ;
            if ( subRoute.PreviousSubRoute is not { } prevSubRoute ) return endPoint ;

            subRoute = prevSubRoute ;
          }
          else {
            var endPoint = subRoute.Segments.FirstOrDefault()?.ToEndPoint.ToEndPointOverSubRoute( document ) ;
            if ( endPoint is not PassPointEndPoint ) return endPoint ;
            if ( subRoute.NextSubRoute is not { } nextSubRoute ) return endPoint ;

            subRoute = nextSubRoute ;
          }
        }
      }
    }

    public static Connector? GetInOutConnector( UIDocument uiDocument, Element eleConn, string message, IPickResult? firstPick, AddInType addInType )
    {
      var document = uiDocument.Document ;

      var pickedObject = uiDocument.Selection.PickObject( ObjectType.Element, FamilyInstanceWithInOutConnectorFilter.GetInstance( addInType ), message ) ;

      var element = document.GetElement( pickedObject?.ElementId ) ;
      if ( null == element ) return null ;

      var connId = element.GetPropertyInt( RoutingFamilyLinkedParameter.RouteConnectorRelationIds ) ;

      return eleConn.GetConnectors().FirstOrDefault( conn => conn.Id == connId ) ;
    }

    #region PickResults

    private class ConnectorPickResult : IPickResult
    {
      private readonly Element _element ;
      private readonly Connector _connector ;
      private readonly string? _routeName ;

      public SubRoute? SubRoute => null ;
      public EndPointKey? EndPointOverSubRoute => null ;
      public Element PickedElement => _element ;
      public Connector? PickedConnector => _connector ;

      public XYZ GetOrigin() => _connector.Origin ;
      public XYZ? GetMEPCurveDirection( bool isFrom ) => _connector.CoordinateSystem.BasisZ ;
      public Element GetOriginElement() => PickedElement ;

      public ConnectorPickResult( Element element, Connector connector )
      {
        _element = element ;
        _connector = connector ;
        _routeName = element.GetRouteName() ;
      }

      public IEnumerable<ElementId> GetAllRelatedElements()
      {
        yield return _element.Id ;
      }
      public ElementId GetLevelId() => _element.GetLevelId() ;

      public bool IsCompatibleTo( Connector connector ) => _connector.IsCompatibleTo( connector ) ;
      public bool IsCompatibleTo( Element element ) => null == _routeName || _routeName != element.GetRouteName() ;
    }

    private class SubRoutePickResult : IPickResult
    {
      private readonly MEPSystemPipeSpec? _spec ;
      private readonly Element _pickedElement ;
      private readonly SubRoute _subRoute ;
      private readonly XYZ _pickPosition ;
      private readonly HashSet<string> _relatedRouteNames ;
      private (Element, XYZ)? _pickPositionOnCenterline ;

      public SubRoute? SubRoute => _subRoute ;
      public EndPointKey? EndPointOverSubRoute { get ; }
      public Element PickedElement => _pickedElement ;
      public Connector? PickedConnector => null ;

      public XYZ GetOrigin() => (_pickPositionOnCenterline ??= GetPickPositionOnCenterline()).Item2 ;
      public XYZ? GetMEPCurveDirection( bool isFrom ) => null ;
      public Element GetOriginElement() => (_pickPositionOnCenterline ??= GetPickPositionOnCenterline()).Item1 ;

      private (Element, XYZ) GetPickPositionOnCenterline()
      {
        if ( null == _spec ) throw new InvalidOperationException() ;

        var halfRequiredLength = GetHalfRequiredLength() ;
        if ( GetNearestMEPCurve( halfRequiredLength * 2 ) is not { } mepCurve ) return ( _pickedElement, GetCenter( _pickedElement ) ) ;

        var connectorPositions = mepCurve.GetConnectors().Where( c => c.IsAnyEnd() ).Select( c => c.Origin ).ToArray() ;
        if ( connectorPositions.Length < 2 ) return ( mepCurve, GetCenter( mepCurve ) ) ;
        XYZ o = connectorPositions[ 0 ], v = connectorPositions[ 1 ] - o ;
        var len = v.GetLength() ;
        v = v.Normalize() ;
        var t = Math.Max( halfRequiredLength, Math.Min( ( _pickPosition - o ).DotProduct( v ), len - halfRequiredLength ) ) ;
        return ( mepCurve, o + t * v ) ;
      }

      private double GetHalfRequiredLength()
      {
        return _subRoute.Route.Document.Application.ShortCurveTolerance ;
      }

      private MEPCurve? GetNearestMEPCurve( double requiredLength )
      {
        var pickedMEPCurve = _pickedElement as MEPCurve ;
        if ( null != pickedMEPCurve && HasEnoughLength( pickedMEPCurve, requiredLength ) ) return pickedMEPCurve ;
        
        var nearestDistance = double.PositiveInfinity ;
        MEPCurve? nearestMEPCurve = null ;
        var subRouteInfo = _pickedElement.GetSubRouteInfo() ;

        foreach ( var connector in _pickedElement.GetConnectors().Where( c => c.IsAnyEnd() ) ) {
          var distance = connector.Origin.DistanceTo( _pickPosition ) ;
          if ( nearestDistance <= distance ) continue ;
          if ( GetConnectingMEPCurve( connector, subRouteInfo, requiredLength ) is not { } mepCurve ) continue ;

          nearestDistance = distance ;
          nearestMEPCurve = mepCurve ;
        }

        return nearestMEPCurve ;

        static MEPCurve? GetConnectingMEPCurve( Connector connector, SubRouteInfo? subRouteInfo, double requiredLength )
        {
          foreach ( var c in connector.GetConnectedConnectors() ) {
            var nextSubRouteInfo = subRouteInfo ;
            if ( c.Owner is MEPCurve mepCurve ) {
              var newSubRouteInfo = mepCurve.GetSubRouteInfo() ;
              if ( nextSubRouteInfo != newSubRouteInfo ) continue ;
              if ( HasEnoughLength( mepCurve, requiredLength ) ) return mepCurve ;

              nextSubRouteInfo = newSubRouteInfo ;
            }

            if ( c.GetOtherConnectorsInOwner().Select( c2 => GetConnectingMEPCurve( c2, nextSubRouteInfo, requiredLength ) ).FirstOrDefault( curve => null != curve ) is { } nextCurve ) return nextCurve ;
          }
          return null ;
        }

        static bool HasEnoughLength( MEPCurve mepCurve, double requiredLength )
        {
          if ( mepCurve.GetRoutingConnectors( true ).FirstOrDefault() is not { } fromConnector ) return false ;
          if ( mepCurve.GetRoutingConnectors( false ).FirstOrDefault() is not { } toConnector ) return false ;

          return ( requiredLength < toConnector.Origin.DistanceTo( fromConnector.Origin ) ) ;
        }
      }

      private SubRoutePickResult( MEPSystemPipeSpec? spec, Element element, SubRoute subRoute, XYZ pickPosition )
      {
        _spec = spec ;
        _pickedElement = element ;
        _subRoute = subRoute ;
        _pickPosition = pickPosition ;
        _relatedRouteNames = _subRoute.Route.GetAllRelatedBranches().Select( route => route.RouteName ).ToHashSet() ;
      }

      private SubRoutePickResult( SubRoutePickResult baseResult, EndPointKey endPointOverSubRoute )
        : this( baseResult._spec, baseResult._pickedElement, baseResult._subRoute, baseResult._pickPosition )
      {
        EndPointOverSubRoute = endPointOverSubRoute ;
        _pickPositionOnCenterline = baseResult._pickPositionOnCenterline ;
      }

      public SubRoutePickResult ApplyEndPointOverSubRoute( EndPointKey endPointOverSubRoute )
      {
        return new SubRoutePickResult( this, endPointOverSubRoute ) ;
      }

      public IEnumerable<ElementId> GetAllRelatedElements()
      {
        return _pickedElement.Document.GetAllElementsOfSubRoute<Element>( _subRoute.Route.RouteName, _subRoute.SubRouteIndex ).Select( e => e.Id ) ;
      }
      public ElementId GetLevelId() => _pickedElement.GetLevelId() ;

      public static SubRoutePickResult? Create( Element element, XYZ pickPosition )
      {
        return Create( (MEPSystemPipeSpec)null!, element, pickPosition ) ;
      }

      public static SubRoutePickResult? Create( RoutingExecutor routingExecutor, Element element, XYZ pickPosition )
      {
        if ( element.GetSubRouteInfo() is not { } subRouteInfo ) return null ;
        if ( RouteCache.Get( element.Document ).GetSubRoute( subRouteInfo ) is not { } subRoute ) return null ;

        return new SubRoutePickResult( routingExecutor.GetMEPSystemPipeSpec( subRoute ), element, subRoute, pickPosition ) ;
      }

      public static SubRoutePickResult? Create( MEPSystemPipeSpec pipeSpec, Element element, XYZ pickPosition )
      {
        if ( element.GetSubRouteInfo() is not { } subRouteInfo ) return null ;
        if ( RouteCache.Get( element.Document ).GetSubRoute( subRouteInfo ) is not { } subRoute ) return null ;

        return new SubRoutePickResult( pipeSpec, element, subRoute, pickPosition ) ;
      }

      public bool IsCompatibleTo( Connector connector ) => _subRoute.GetReferenceConnector().IsCompatibleTo( connector ) ;

      public bool IsCompatibleTo( Element element )
      {
        return ( element.GetRouteName() is not { } routeName ) || ( false == _relatedRouteNames.Contains( routeName ) ) ;
      }
    }

    private class PassPointPickResult : IPickResult
    {
      private readonly Element _element ;
      private readonly Route? _route ;

      public SubRoute? SubRoute => null ;
      public EndPointKey? EndPointOverSubRoute => null ;
      public Element PickedElement => _element ;
      public Connector? PickedConnector => null ;

      public XYZ GetOrigin() => GetCenter( _element ) ;
      public XYZ? GetMEPCurveDirection( bool isFrom ) => null ;
      public Element GetOriginElement() => PickedElement ;

      private PassPointPickResult( Element element )
      {
        _element = element ;

        if ( element.GetRouteName() is { } routeName ) {
          RouteCache.Get( element.Document ).TryGetValue( routeName, out _route ) ;
        }
      }

      public IEnumerable<ElementId> GetAllRelatedElements()
      {
        return _element.Document.GetAllElementsOfPassPoint( _element.GetPassPointId() ?? _element.Id.IntegerValue ).Select( e => e.Id ) ;
      }
      public ElementId GetLevelId() => _element.GetLevelId() ;

      public static IPickResult? Create( Element element )
      {
        if ( false == element.IsPassPoint() ) return null ;

        if ( element.GetPassPointId() is { } i && i != element.Id.IntegerValue ) {
          element = element.Document.GetElement( new ElementId( i ) ) ;
          if ( null == element || false == element.IsPassPoint() ) return null ;
        }

        return new PassPointPickResult( element ) ;
      }

      public bool IsCompatibleTo( Connector connector ) => null == _route || _route.GetReferenceConnector().IsCompatibleTo( connector ) ;
      public bool IsCompatibleTo( Element element ) => null == _route || _route.RouteName != element.GetRouteName() ;
    }

    private class OriginPickResult : IPickResult
    {
      private readonly Element _element ;
      private readonly AddInType _addInType ;

      public SubRoute? SubRoute => null ;
      public EndPointKey? EndPointOverSubRoute => null ;
      public Element PickedElement => _element ;
      public Connector? PickedConnector => null ;

      public XYZ GetOrigin() => GetCenter( _element ) ;
      public XYZ? GetMEPCurveDirection( bool isFrom ) => XYZ.Zero ;
      public Element GetOriginElement() => PickedElement ;

      public OriginPickResult( Element element, AddInType addInType )
      {
        _element = element ;
        _addInType = addInType ;
      }

      public IEnumerable<ElementId> GetAllRelatedElements()
      {
        yield return _element.Id ;
      }
      public ElementId GetLevelId() => _element.GetLevelId() ;

      public bool IsCompatibleTo( Connector connector ) => true ;

      public bool IsCompatibleTo( Element element ) => element.GetConnectors().Any( c => IsTargetConnector( c, _addInType ) ) ;
    }

    private static XYZ GetCenter( Element element )
    {
      return element switch
      {
        MEPCurve curve => GetCenter( curve ),
        Instance instance => instance.GetTotalTransform().Origin,
        Room room => ( (LocationPoint) room.Location ).Point,
        _ => throw new System.InvalidOperationException(),
      } ;
    }

    private static XYZ GetCenter( MEPCurve curve )
    {
      double minX = +double.MaxValue, minY = -double.MaxValue, minZ = +double.MaxValue ;
      double maxX = -double.MaxValue, maxY = +double.MaxValue, maxZ = -double.MaxValue ;

      foreach ( var c in curve.GetConnectors().Where( c => c.IsAnyEnd() ) ) {
        var (x, y, z) = c.Origin ;

        if ( x < minX ) minX = x ;
        if ( maxX < x ) maxX = x ;
        if ( y < minY ) minY = y ;
        if ( maxY < y ) maxY = y ;
        if ( z < minZ ) minZ = z ;
        if ( maxZ < z ) maxZ = z ;
      }

      return new XYZ( ( minX + maxX ) * 0.5, ( minY + maxY ) * 0.5, ( minZ + maxZ ) * 0.5 ) ;
    }

    #endregion

    private static (bool Result, Connector? Connector) FindConnector( UIDocument uiDocument, Element element, string message, Connector? firstConnector, AddInType addInType )
    {
      if ( element.IsAutoRoutingGeneratedElement() ) {
        return GetEndOfRouting( element, ( null == firstConnector ) ) ;
      }
      else {
        return CreateConnectorInOutFamily( uiDocument, element, message, firstConnector, addInType ) ;
      }
    }

    private static (bool Result, Connector? Connector) GetEndOfRouting( Element element, bool fromConnector )
    {
      var routeName = element.GetRouteName() ;
      if ( null == routeName ) return ( false, null ) ;

      var connector = element.Document.CollectRoutingEndPointConnectors( routeName, fromConnector ).FirstOrDefault() ;
      return ( ( null != connector ), connector ) ;
    }

    private static (bool Result, Connector? Connector) CreateConnectorInOutFamily( UIDocument uiDocument, Element element, string message, Connector? firstConnector, AddInType addInType )
    {
      using var fitter = new TempZoomToFit( uiDocument ) ;

      uiDocument.SetSelection( element ) ;
      fitter.ZoomToFit() ;

      var sv = new CreateConnector( uiDocument, element, firstConnector, addInType ) ;

      uiDocument.ClearSelection() ;
      return ( true, sv.GetPickedConnector() ) ;
    }

    private static (bool Result, Connector? Connector) SelectFromDialog( UIDocument uiDocument, Element element, string message, Connector? firstConnector, AddInType addInType )
    {
      using var fitter = new TempZoomToFit( uiDocument ) ;

      uiDocument.SetSelection( element ) ;
      fitter.ZoomToFit() ;

      var sv = new SelectConnector( element, firstConnector, addInType ) { Title = message } ;
      sv.ShowDialog() ;

      uiDocument.ClearSelection() ;

      if ( true != sv.DialogResult ) return ( false, null ) ;

      return ( true, sv.GetSelectedConnector() ) ;
    }

    public static bool IsTargetConnector( Connector conn, AddInType addInType ) => conn.IsAnyEnd() && IsTargetDomain( conn.Domain, addInType ) ;

    public static bool IsTargetDomain( Domain domain, AddInType addInType )
    {
      return domain switch
      {
        Domain.DomainPiping => ( addInType == AddInType.Mechanical ),
        Domain.DomainHvac => ( addInType == AddInType.Mechanical ),
        Domain.DomainElectrical => ( addInType == AddInType.Electrical ),
        Domain.DomainCableTrayConduit => ( addInType == AddInType.Electrical ),
        _ => false,
      } ;
    }

    private class FamilyInstanceWithConnectorFilter : ISelectionFilter
    {
      private static readonly Dictionary<AddInType, ISelectionFilter> _dic = new() ;

      public static ISelectionFilter GetInstance( AddInType addInType )
      {
        if ( false == _dic.TryGetValue( addInType, out var filter ) ) {
          filter = new FamilyInstanceWithConnectorFilter( addInType, true ) ;
          _dic.Add( addInType, filter ) ;
        }

        return filter ;
      }

      private readonly AddInType _addInType ;

      private bool AllowRoute { get ; }

      protected FamilyInstanceWithConnectorFilter( AddInType addInType, bool allowRoute )
      {
        _addInType = addInType ;
        AllowRoute = allowRoute ;
      }

      public bool AllowElement( Element elem )
      {
        return IsRoutableForConnector( elem ) || IsRoutableForCenter( elem ) || IsRoutableForRoom( elem ) ;
      }

      private bool IsRoutableForConnector( Element elem )
      {
        return elem.GetConnectors().Any( IsTargetConnector ) && IsRoutableElement( elem ) ;
      }

      private static bool IsRoutableForCenter( Element elem )
      {
        return ( elem is FamilyInstance fi ) && ( false == fi.IsPassPoint() ) && ( false == elem.IsAutoRoutingGeneratedElement() ) ;
      }

      private static bool IsRoutableForRoom( Element elem )
      {
        return ( elem is Room ) ;
      }

      protected virtual bool IsRoutableElement( Element elem )
      {
        return elem switch
        {
          MEPCurve => AllowRoute,
          FamilyInstance fi => IsEquipment( fi ) || ( AllowRoute && elem.IsAutoRoutingGeneratedElement() ),
          _ => false,
        } && PointOnRouteFilters.RepresentativeElement( elem ) ;
      }

      private static bool IsEquipment( FamilyInstance fi )
      {
        if ( fi.IsFittingElement() ) return false ;
        if ( fi.IsPassPoint() ) return false ;
        return true ;
      }

      protected virtual bool IsTargetConnector( Connector connector )
      {
        return ConnectorPicker.IsTargetConnector( connector, _addInType ) ;
      }

      public bool AllowReference( Reference reference, XYZ position )
      {
        return false ;
      }
    }

    private class FamilyInstanceWithInOutConnectorFilter : ISelectionFilter
    {
      private static readonly Dictionary<AddInType, ISelectionFilter> _dic = new() ;

      public static ISelectionFilter GetInstance( AddInType addInType )
      {
        if ( false == _dic.TryGetValue( addInType, out var filter ) ) {
          filter = new FamilyInstanceWithInOutConnectorFilter( addInType ) ;
          _dic.Add( addInType, filter ) ;
        }

        return filter ;
      }

      private readonly AddInType _addInType ;

      private FamilyInstanceWithInOutConnectorFilter( AddInType addInType )
      {
        _addInType = addInType ;
      }

      public bool AllowElement( Element elem )
      {
        return IsRoutableForCenter( elem ) ;
      }


      private static bool IsRoutableForCenter( Element elem )
      {
        return ( elem is FamilyInstance fi ) && ( true == fi.IsConnectorPoint() ) ;
      }

      protected virtual bool IsRoutableElement( Element elem )
      {
        return elem switch
        {
          MEPCurve => true,
          FamilyInstance fi => IsEquipment( fi ) || elem.IsAutoRoutingGeneratedElement(),
          _ => false,
        } ;
      }

      private static bool IsEquipment( FamilyInstance fi )
      {
        if ( fi.IsFittingElement() ) return false ;
        if ( fi.IsPassPoint() ) return false ;
        return true ;
      }

      protected virtual bool IsTargetConnector( Connector connector )
      {
        return ConnectorPicker.IsTargetConnector( connector, _addInType ) ;
      }

      public bool AllowReference( Reference reference, XYZ position )
      {
        return false ;
      }
    }

    private class FamilyInstanceCompatibleToTargetConnectorFilter : FamilyInstanceWithConnectorFilter
    {
      private readonly IPickResult _compatibleResult ;

      public FamilyInstanceCompatibleToTargetConnectorFilter( IPickResult compatibleResult, AddInType addInType ) : base( addInType, AllowsPickRoute( compatibleResult, addInType ) )
      {
        _compatibleResult = compatibleResult ;
      }

      private static bool AllowsPickRoute( IPickResult compatibleResult, AddInType addInType )
      {
        if ( AddInType.Electrical == addInType ) return true ;
        if ( null == compatibleResult.SubRoute ) return true ;

        return false ;
      }

      protected override bool IsTargetConnector( Connector connector )
      {
        return base.IsTargetConnector( connector ) && _compatibleResult.IsCompatibleTo( connector ) ;
      }

      protected override bool IsRoutableElement( Element elem )
      {
        return base.IsRoutableElement( elem ) && _compatibleResult.IsCompatibleTo( elem ) ;
      }
    }
  }
}