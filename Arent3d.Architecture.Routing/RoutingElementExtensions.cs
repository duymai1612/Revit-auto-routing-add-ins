using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;
using Autodesk.Revit.DB.Mechanical ;
using Autodesk.Revit.DB.Plumbing ;
using Autodesk.Revit.DB.Structure ;

namespace Arent3d.Architecture.Routing
{
  public static class RoutingElementExtensions
  {
    #region Initializations

    /// <summary>
    /// Confirms whether families and parameters used for routing application are loaded.
    /// </summary>
    /// <param name="document"></param>
    /// <returns>True if all families and parameters are loaded.</returns>
    public static bool RoutingSettingsAreInitialized( this Document document )
    {
      return document.AllRoutingFamiliesAreLoaded() && document.AllRoutingParametersAreRegistered() ;
    }

    /// <summary>
    /// Setup all families and parameters used for routing application.
    /// </summary>
    /// <param name="document"></param>
    public static bool SetupRoutingFamiliesAndParameters( this Document document )
    {
      document.MakeCertainAllRoutingFamilies() ;
      document.MakeCertainAllRoutingParameters() ;
      AddArentConduitType( document ) ;

      //Add connector type value
      var connectorOneSide = document.GetAllElements<FamilyInstance>().OfCategory( BuiltInCategorySets.Connectors ) ;
      foreach ( var connector in connectorOneSide ) {
        connector.SetConnectorFamilyType( ConnectorFamilyType.Sensor ) ;
      }

      return document.RoutingSettingsAreInitialized() ;
    }

    /// <summary>
    /// Setup all families and parameters used for routing application.
    /// </summary>
    /// <param name="document"></param>
    public static void UnsetupRoutingFamiliesAndParameters( this Document document )
    {
      EraseArentConduitType( document ) ;
      document.EraseAllRoutingFamilies() ;
      document.UnloadAllRoutingParameters() ;
    }

    private static string GetElbowTypeName( Document document )
    {
      return "Routing.Revit.DummyConduit.ElbowTypeName".GetDocumentStringByKeyOrDefault( document, "M_電線管エルボ - 鉄鋼" ) ;
    }
    private static string GetConduitTypeName( Document document )
    {
      return "Routing.Revit.DummyConduit.ConduitTypeName".GetDocumentStringByKeyOrDefault( document, "Arent電線" ) ;
    }

    private static double GetConduitTypeNominalDiameter( Document document )
    {
      return ( 1.0 ).MillimetersToRevitUnits() ;
    }
    
    private static void AddArentConduitType( Document document )
    {
      var elbowTypeName = GetElbowTypeName( document ) ;
      var conduitTypeName = GetConduitTypeName( document ) ;
      var sizes = ConduitSizeSettings.GetConduitSizeSettings( document ) ;
      var standards = document.GetStandardTypes().ToList() ;
      if ( ! standards.Contains( conduitTypeName ) ) {
        var conduitStandard = sizes.CreateConduitStandardTypeFromExisingStandardType( document, conduitTypeName, standards.Last() ) ;
        if ( conduitStandard ) {
          ConduitSize sizeInfo = new ConduitSize( GetConduitTypeNominalDiameter( document ), ( 0.8 ).MillimetersToRevitUnits(), ( 1.2 ).MillimetersToRevitUnits(), ( 16.0 ).MillimetersToRevitUnits(), true, true ) ;
          sizes.AddSize( standards.Last(), sizeInfo ) ;
          sizes.AddSize( conduitTypeName, sizeInfo ) ;
        }
      }

      var elbowCurveType = document.GetAllElements<FamilySymbol>().OfCategory( BuiltInCategory.OST_ConduitFitting ).FirstOrDefault( x => x.FamilyName == elbowTypeName ) ;

      // Get type by "Standard" parameter
      var curveTypes = document.GetAllElements<ConduitType>().Where( c => c.get_Parameter( BuiltInParameter.CONDUIT_STANDARD_TYPE_PARAM ).AsValueString() == standards.Last() ).OfType<MEPCurveType>().ToList() ;
      foreach ( var curveType in curveTypes ) {
        var arentConduitType = document.GetAllElements<ConduitType>().FirstOrDefault( c => c.Name == conduitTypeName && c.FamilyName == curveType.FamilyName ) ;
        if ( arentConduitType != null ) continue ;
        var arentCurveType = curveType.Duplicate( conduitTypeName ) ;
        
        // Change "Bend" value
        var arentBend = arentCurveType.get_Parameter( BuiltInParameter.RBS_CURVETYPE_DEFAULT_BEND_PARAM ).Element as ConduitType ;
        arentBend!.Elbow = elbowCurveType ;
      }
    }

    private static void EraseArentConduitType( Document document )
    {
      var elbowTypeName = GetElbowTypeName( document ) ;
      var conduitTypeName = GetConduitTypeName( document ) ;
      var sizes = ConduitSizeSettings.GetConduitSizeSettings( document ) ;
      var standards = document.GetStandardTypes().ToList() ;
      if ( standards.Contains( conduitTypeName ) ) {
        // todo check existence
        sizes.RemoveSize( conduitTypeName, GetConduitTypeNominalDiameter( document ) ) ;
        sizes.RemoveSize( standards.Last(), GetConduitTypeNominalDiameter( document ) ) ;
      }

      var curveTypes = document.GetAllElements<ConduitType>().Where( c => c.get_Parameter( BuiltInParameter.CONDUIT_STANDARD_TYPE_PARAM ).AsValueString() == standards.Last() ).OfType<MEPCurveType>().ToList() ;
      var eraseIds = curveTypes.Where( curveType => curveType.Name == conduitTypeName ).Select( curveType => curveType.Id ).ToList() ;
      document.Delete( eraseIds ) ;
    }

    #endregion

    #region Connectors (General)

    public static Connector? FindConnector( this Document document, int elementId, int connectorId )
    {
      return document.GetElement( new ElementId( elementId ) ).GetConnectorManager()?.Lookup( connectorId ) ;
    }

    public static ConnectorManager? GetConnectorManager( this Element elm )
    {
      return elm switch
      {
        FamilyInstance fi => fi.MEPModel?.ConnectorManager,
        MEPSystem sys => sys.ConnectorManager,
        MEPCurve crv => crv.ConnectorManager,
        _ => null,
      } ;
    }

    public static IEnumerable<Connector> GetConnectors( this Element elm )
    {
      if ( ! ( elm.GetConnectorManager()?.Connectors is { } connectorSet ) ) return Array.Empty<Connector>() ;

      return connectorSet.OfType<Connector>() ;
    }

    public static ConnectorSet ToConnectorSet( this IEnumerable<Connector> connectors )
    {
      var connectorSet = new ConnectorSet() ;

      foreach ( var connector in connectors ) {
        connectorSet.Insert( connector ) ;
      }

      return connectorSet ;
    }

    /// <summary>
    /// Returns connected connectors.
    /// </summary>
    /// <param name="connector"></param>
    /// <returns></returns>
    public static IEnumerable<Connector> GetConnectedConnectors( this Connector connector )
    {
      var id = connector.Owner.Id ;
      return connector.AllRefs.OfType<Connector>().Where( c => c.Owner.Id != id ) ;
    }

    public static IEnumerable<Connector> OfEnd( this IEnumerable<Connector> connectors )
    {
      return connectors.Where( c => c.IsAnyEnd() ) ;
    }

    public static bool IsAnyEnd( this Connector conn )
    {
      return 0 != ( (int)conn.ConnectorType & (int)ConnectorType.AnyEnd ) ;
    }

    public static IEnumerable<Connector> GetOtherConnectorsInOwner( this Connector connector )
    {
      var id = connector.Id ;
      var manager = connector.ConnectorManager ;
      if ( null == manager ) return Array.Empty<Connector>() ;

      return manager.Connectors.OfType<Connector>().Where( c => c.Id != id ) ;
    }

    public static bool HasCompatibleSystemType( this Connector connector1, Connector connector2 )
    {
      if ( connector1.Domain != connector2.Domain ) return false ;

      var classification1 = connector1.GetSystemType() ;
      var classification2 = connector2.GetSystemType() ;
      if ( classification1.IsCompatibleTo( classification2 ) ) return true ;

      return false ;
    }


    public static MEPSystemClassification GetSystemType( this Connector conn )
    {
      return conn.Domain switch
      {
        Domain.DomainPiping => (MEPSystemClassification)conn.PipeSystemType,
        Domain.DomainHvac => (MEPSystemClassification)conn.DuctSystemType,
        Domain.DomainElectrical => (MEPSystemClassification)conn.ElectricalSystemType,
        Domain.DomainCableTrayConduit => MEPSystemClassification.CableTrayConduit,
        _ => MEPSystemClassification.UndefinedSystemClassification,
      } ;
    }

    public static bool HasCompatibleSystemType( this Connector connector, MEPSystemClassification systemClassification )
    {
      if ( systemClassification == MEPSystemClassification.Global || systemClassification == MEPSystemClassification.Fitting ) {
        return true ;
      }

      return systemClassification.IsCompatibleTo( connector.GetSystemType() ) ;
    }

    private static bool IsCompatibleTo( this MEPSystemClassification systemClassification1, MEPSystemClassification systemClassification2 )
    {
      if ( systemClassification1 == MEPSystemClassification.UndefinedSystemClassification ) return true ;
      if ( systemClassification2 == MEPSystemClassification.UndefinedSystemClassification ) return true ;

      if ( systemClassification1 == MEPSystemClassification.PowerCircuit && IsCompatibleToPowerCircuit( systemClassification2 ) ) return true ;
      if ( systemClassification2 == MEPSystemClassification.PowerCircuit && IsCompatibleToPowerCircuit( systemClassification1 ) ) return true ;

      return ( systemClassification1 == systemClassification2 ) ;
    }

    private static bool IsCompatibleToPowerCircuit( MEPSystemClassification systemClassification )
    {
      return ( systemClassification == MEPSystemClassification.PowerBalanced || systemClassification == MEPSystemClassification.PowerUnBalanced ) ;
    }

    public static Connector GetTopConnectorOfConnectorFamily( this FamilyInstance elm )
    {
      var topItem = elm.GetConnectors().MaxBy( conn => conn.Origin.Z ) ;
      return topItem ?? elm.GetConnectors().First() ;
    }

    #endregion

    #region Connectors (Routing)

    public static bool IsCompatibleTo( this Connector conn1, Connector conn2 )
    {
      return ( conn1.ConnectorType == conn2.ConnectorType ) && ( conn1.Domain == conn2.Domain ) && conn1.HasCompatibleShape( conn2 ) ;
    }

    public static bool HasCompatibleShape( this IConnector conn1, IConnector conn2 )
    {
      return ( conn1.Shape == conn2.Shape ) ;
    }

    public static bool HasSameShapeAndParameters( this IConnector conn1, IConnector conn2 )
    {
      if ( conn1.Shape != conn2.Shape ) return false ;

      // Concrete shape parameter can be different
      return conn1.Shape switch
      {
        ConnectorProfileType.Oval => HasSameOvalShape( conn1, conn2 ),
        ConnectorProfileType.Round => HasSameRoundShape( conn1, conn2 ),
        ConnectorProfileType.Rectangular => HasSameRectangularShape( conn1, conn2 ),
        _ => false,
      } ;
    }

    private static bool HasSameOvalShape( IConnector conn1, IConnector conn2 )
    {
      var document = GetDocument( conn1 ) ?? GetDocument( conn2 ) ?? throw new InvalidOperationException() ;
      var tole = document.Application.VertexTolerance ;
      return Math.Abs( conn1.Radius - conn2.Radius ) < tole ;
    }

    private static bool HasSameRoundShape( IConnector conn1, IConnector conn2 )
    {
      var document = GetDocument( conn1 ) ?? GetDocument( conn2 ) ?? throw new InvalidOperationException() ;
      var tole = document.Application.VertexTolerance ;
      return Math.Abs( conn1.Radius - conn2.Radius ) < tole ;
    }

    private static bool HasSameRectangularShape( IConnector conn1, IConnector conn2 )
    {
      var document = GetDocument( conn1 ) ?? GetDocument( conn2 ) ?? throw new InvalidOperationException() ;
      var tole = document.Application.VertexTolerance ;
      return Math.Abs( conn1.Width - conn2.Width ) < tole && Math.Abs( conn1.Height - conn2.Height ) < tole ;
    }

    private static Document? GetDocument( IConnector conn )
    {
      return conn switch
      {
        Connector c => c.Owner?.Document,
        Element ce => ce.Document,
        _ => null,
      } ;
    }

    #endregion

    #region Pass Points

    public static FamilyInstance? FindPassPointElement( this Document document, int elementId )
    {
      var instance = document.GetElementById<FamilyInstance>( elementId ) ;
      if ( null == instance ) return null ;

      if ( instance.Symbol.Id != document.GetFamilySymbols( RoutingFamilyType.PassPoint ).FirstOrDefault().GetValidId() ) {
        // Family instance is not a pass point.
        return null ;
      }

      return instance ;
    }

    public static void SetPassPointConnectors( this Element element, IReadOnlyCollection<Connector> fromConnectors, IReadOnlyCollection<Connector> toConnectors )
    {
      element.SetProperty( PassPointParameter.PassPointNextToFromSideConnectorIds, string.Join( "|", fromConnectors.Select( ConnectorEndPoint.BuildParameterString ) ) ) ;
      element.SetProperty( PassPointParameter.PassPointNextToToSideConnectorIds, string.Join( "|", toConnectors.Select( ConnectorEndPoint.BuildParameterString ) ) ) ;
    }

    private static readonly char[] PassPointConnectorSeparator = { '|' } ;

    public static IEnumerable<IEndPoint> GetPassPointConnectors( this Element element, bool isFrom )
    {
      var parameter = isFrom ? PassPointParameter.PassPointNextToFromSideConnectorIds : PassPointParameter.PassPointNextToToSideConnectorIds ;
      if ( false == element.TryGetProperty( parameter, out string? str ) ) return Array.Empty<IEndPoint>() ;
      if ( string.IsNullOrEmpty( str ) ) return Array.Empty<IEndPoint>() ;

      var document = element.Document ;
      return str!.Split( PassPointConnectorSeparator, StringSplitOptions.RemoveEmptyEntries ).Select( s => ConnectorEndPoint.ParseParameterString( document, s ) ).NonNull() ;
    }

    public static bool IsPassPoint( this Element element )
    {
      return element is FamilyInstance fi && fi.IsPassPoint() ;
    }

    public static bool IsPassPoint( this FamilyInstance element )
    {
      return element.IsFamilyInstanceOf( RoutingFamilyType.PassPoint ) || element.HasParameter( RoutingParameter.RelatedPassPointId ) ;
    }

    public static bool IsConnectorPoint( this FamilyInstance element )
    {
      return element.IsFamilyInstanceOfAny( RoutingFamilyType.ConnectorInPoint, RoutingFamilyType.ConnectorOutPoint, RoutingFamilyType.ConnectorPoint, RoutingFamilyType.TerminatePoint ) ;
    }

    public static int? GetPassPointId( this Element element )
    {
      if ( element is not FamilyInstance fi ) return null ;

      if ( fi.IsFamilyInstanceOf( RoutingFamilyType.PassPoint ) ) return fi.Id.IntegerValue ;
      if ( element.TryGetProperty( RoutingParameter.RelatedPassPointId, out int id ) ) return id ;
      return null ;
    }

    public static FamilyInstance AddPassPoint( this Document document, string routeName, XYZ position, XYZ direction, double? radius, ElementId levelId )
    {
      var instance = document.CreateFamilyInstance( RoutingFamilyType.PassPoint, position, direction, StructuralType.NonStructural, true, document.GetElementById<Level>( levelId ) ) ;
      if ( radius.HasValue ) {
        instance.LookupParameter( "Arent-RoundDuct-Diameter" ).Set( radius.Value * 2.0 ) ;
      }

      instance.SetProperty( RoutingParameter.RouteName, routeName ) ;

      return instance ;
    }

    public static FamilyInstance AddConnectorFamily( this Document document, Connector conn, string routeName, FlowDirectionType directionType, XYZ position, XYZ direction, double? radius )
    {
      var routingFamilyType = directionType switch
      {
        FlowDirectionType.In => RoutingFamilyType.ConnectorInPoint,
        FlowDirectionType.Out => RoutingFamilyType.ConnectorOutPoint,
        _ => RoutingFamilyType.ConnectorPoint,
      } ;

      var instance = document.CreateFamilyInstance( routingFamilyType, position, StructuralType.NonStructural, true, document.GetElementById<Level>( conn.Owner.LevelId ) ) ;
      var id = conn.Id ;

      instance.SetProperty( RoutingFamilyLinkedParameter.RouteConnectorRelationIds, id ) ;


      var elevationAngle = Math.Atan2( direction.Z, Math.Sqrt( direction.X * direction.X + direction.Y * direction.Y ) ) ;
      Color colorIn = new Autodesk.Revit.DB.Color( (byte)255, (byte)0, (byte)0 ) ;
      Color colorOut = new Autodesk.Revit.DB.Color( (byte)0, (byte)0, (byte)255 ) ;
      OverrideGraphicSettings ogsIn = new OverrideGraphicSettings() ;
      OverrideGraphicSettings ogsOut = new OverrideGraphicSettings() ;
      ogsIn.SetProjectionLineColor( colorIn ) ;
      ogsOut.SetProjectionLineColor( colorOut ) ;

      if ( directionType == FlowDirectionType.Out ) {
        //Out
        document.ActiveView.SetElementOverrides( instance.Id, ogsIn ) ;
        if ( conn.CoordinateSystem.BasisX.Y > 0 ) {
          var rotationAngle = Math.Atan2( -direction.Y, direction.X ) ;

          ElementTransformUtils.RotateElement( document, instance.Id, Line.CreateBound( position, position + XYZ.BasisZ ), rotationAngle ) ;
        }
        else {
          var rotationAngle = Math.Atan2( direction.Y, direction.X ) ;
          ElementTransformUtils.RotateElement( document, instance.Id, Line.CreateBound( position, position + XYZ.BasisZ ), rotationAngle ) ;
        }
      }
      else if ( directionType == FlowDirectionType.In ) {
        //In
        document.ActiveView.SetElementOverrides( instance.Id, ogsOut ) ;
        if ( conn.CoordinateSystem.BasisX.Y > 0 ) {
          var rotationAngle = Math.Atan2( direction.Y, direction.X ) ;
          ElementTransformUtils.RotateElement( document, instance.Id, Line.CreateBound( position, position + XYZ.BasisZ ), rotationAngle ) ;
        }
        else {
          var rotationAngle = Math.Atan2( -direction.Y, direction.X ) ;
          ElementTransformUtils.RotateElement( document, instance.Id, Line.CreateBound( position, position + XYZ.BasisZ ), rotationAngle ) ;
        }
      }

      instance.SetProperty( RoutingParameter.RouteName, routeName ) ;

      return instance ;
    }

    public static FamilyInstance AddRackGuide( this Document document, XYZ position, Level? level )
    {
      return document.CreateFamilyInstance( RoutingFamilyType.RackGuide, position, StructuralType.NonStructural, true, level ) ;
    }
    public static FamilyInstance AddRackSpace( this Document document, XYZ position, Level level )
    {
      return document.CreateFamilyInstance( RoutingFamilyType.RackSpace, position, StructuralType.NonStructural, true, level ) ;
    }
    public static FamilyInstance AddFASU(this Document document, RoutingFamilyType fasuFamilyType, XYZ position, ElementId levelId)
    {
      return document.CreateFamilyInstance( fasuFamilyType, position, StructuralType.NonStructural, true, document.GetElementById<Level>(levelId));
    }
    public static FamilyInstance AddVAV(this Document document, XYZ position, ElementId levelId)
    {
      return document.CreateFamilyInstance( RoutingFamilyType.TTE_VAV_140, position, StructuralType.NonStructural, true, document.GetElementById<Level>(levelId));
    }
    public static FamilyInstance AddShaft( this Document document, XYZ position, Level? level )
    {
      return document.CreateFamilyInstance( RoutingFamilyType.Shaft, position, StructuralType.NonStructural, true, level ) ;
    }

    public static FamilyInstance AddCornPoint( this Document document, string routeName, XYZ position, Level? level )
    {
      var instance = document.CreateFamilyInstance( RoutingFamilyType.CornPoint, position, StructuralType.NonStructural, true, level ) ;
      instance.SetProperty( RoutingParameter.RouteName, routeName ) ;

      return instance ;
    }

    #endregion

    #region Terminate Points

    public static FamilyInstance? FindTerminatePointElement( this Document document, int elementId )
    {
      var instance = document.GetElementById<FamilyInstance>( elementId ) ;
      if ( null == instance ) return null ;

      if ( instance.Symbol.Id != document.GetFamilySymbols( RoutingFamilyType.TerminatePoint ).FirstOrDefault()?.Id ) {
        // Family instance is not a pass point.
        return null ;
      }

      return instance ;
    }

    public static bool IsTerminatePoint( this Element element )
    {
      return element is FamilyInstance fi && fi.IsTerminatePoint() ;
    }

    public static bool IsTerminatePoint( this FamilyInstance element )
    {
      return element.IsFamilyInstanceOf( RoutingFamilyType.TerminatePoint ) || element.HasParameter( RoutingParameter.RelatedTerminatePointId ) ;
    }

    public static int? GetTerminatePointId( this Element element )
    {
      if ( element is not FamilyInstance fi ) return null ;

      if ( fi.IsFamilyInstanceOf( RoutingFamilyType.TerminatePoint ) ) return fi.Id.IntegerValue ;
      if ( element.TryGetProperty( RoutingParameter.RelatedTerminatePointId, out int id ) ) return id ;
      return null ;
    }

    public static FamilyInstance AddTerminatePoint( this Document document, string routeName, XYZ position, XYZ direction, double? radius, ElementId levelId )
    {
      var instance = document.CreateFamilyInstance( RoutingFamilyType.TerminatePoint, position, direction, StructuralType.NonStructural, true, document.GetElementById<Level>( levelId ) ) ;
      if ( radius.HasValue ) {
        instance.LookupParameter( "Arent-RoundDuct-Diameter" ).Set( radius.Value * 2.0 ) ;
      }

      instance.SetProperty( RoutingParameter.RouteName, routeName ) ;

      return instance ;
    }

    #endregion

    #region Routing (General)

    public static bool IsAutoRoutingGeneratedElement( this Element element )
    {
      return element.IsAutoRoutingGeneratedElementType() && element.HasParameter( RoutingParameter.RouteName ) ;
    }

    public static bool IsAutoRoutingGeneratedElementType( this Element element )
    {
      return element switch
      {
        Duct or Pipe or CableTray => true,
        _ => IsFittingElement( element ),
      } ;
    }

    public static bool IsFittingElement( this Element element )
    {
      var category = element.Category ;
      if ( null == category ) return false ;
      return ( category.CategoryType == CategoryType.Model && IsFittingCategory( category.GetBuiltInCategory() ) ) ;
    }

    private static bool IsFittingCategory( BuiltInCategory category )
    {
      return ( 0 <= Array.IndexOf( BuiltInCategorySets.Fittings, category ) ) ;
    }

    #endregion

    #region Routing (Route Names)

    public static IEnumerable<TElement> GetAllElementsOfRoute<TElement>( this Document document ) where TElement : Element
    {
      var parameterName = document.GetParameterName( RoutingParameter.RouteName ) ;
      if ( null == parameterName ) return Array.Empty<TElement>() ;

      var filter = new ElementParameterFilter( ParameterFilterRuleFactory.CreateSharedParameterApplicableRule( parameterName ) ) ;

      return document.GetAllElementsOfRouteName<TElement>( BuiltInCategorySets.RoutingElements, filter ) ;
    }

    public static IEnumerable<TElement> GetAllElementsOfRouteName<TElement>( this Document document, string routeName ) where TElement : Element
    {
      var parameterName = document.GetParameterName( RoutingParameter.RouteName ) ;
      if ( null == parameterName ) return Array.Empty<TElement>() ;

      var filter = new ElementParameterFilter( ParameterFilterRuleFactory.CreateSharedParameterApplicableRule( parameterName ) ) ;

      return document.GetAllElementsOfRouteName<TElement>( BuiltInCategorySets.RoutingElements, filter ).Where( e => e.GetRouteName() == routeName ) ;
    }

    public static IEnumerable<TElement> GetAllElementsOfRepresentativeRouteName<TElement>( this Document document, string routeName ) where TElement : Element
    {
      var parameterName = document.GetParameterName( RoutingParameter.RepresentativeRouteName ) ;
      if ( null == parameterName ) return Array.Empty<TElement>() ;

      var filter = new ElementParameterFilter( ParameterFilterRuleFactory.CreateSharedParameterApplicableRule( parameterName ) ) ;

      return document.GetAllElementsOfRouteName<TElement>( BuiltInCategorySets.RoutingElements, filter ).Where( e => e.GetRepresentativeRouteName() == routeName ) ;
    }

    public static IEnumerable<TElement> GetAllElementsOfSubRoute<TElement>( this Document document, string routeName, int subRouteIndex ) where TElement : Element
    {
      var routeNameParameterName = document.GetParameterName( RoutingParameter.RouteName ) ;
      if ( null == routeNameParameterName ) return Array.Empty<TElement>() ;

      var subRouteIndexParameterName = document.GetParameterName( RoutingParameter.SubRouteIndex ) ;
      if ( null == subRouteIndexParameterName ) return Array.Empty<TElement>() ;

      var filter = new ElementParameterFilter( new[] { ParameterFilterRuleFactory.CreateSharedParameterApplicableRule( routeNameParameterName ), ParameterFilterRuleFactory.CreateSharedParameterApplicableRule( subRouteIndexParameterName ), } ) ;

      return document.GetAllElementsOfRouteName<TElement>( BuiltInCategorySets.RoutingElements, filter ).Where( e => e.GetRouteName() == routeName ).Where( e => e.GetSubRouteIndex() == subRouteIndex ) ;
    }

    private static IEnumerable<TElement> GetAllElementsOfRouteName<TElement>( this Document document, BuiltInCategory[] builtInCategories, ElementFilter filter ) where TElement : Element
    {
      return document.GetAllElements<Element>().OfCategory( builtInCategories ).OfNotElementType().Where( filter ).OfType<TElement>() ;
    }

    public static IEnumerable<FamilyInstance> GetAllElementsOfPassPoint( this Document document, int passPointId )
    {
      var parameterName = document.GetParameterName( RoutingParameter.RelatedPassPointId ) ;
      if ( null == parameterName ) yield break ;

      var elm = document.GetElementById<FamilyInstance>( passPointId ) ;
      if ( null == elm ) yield break ;
      if ( elm.IsFamilyInstanceOf( RoutingFamilyType.PassPoint ) ) yield return elm ;

      var filter = new ElementParameterFilter( ParameterFilterRuleFactory.CreateSharedParameterApplicableRule( parameterName ) ) ;

      foreach ( var e in document.GetAllElements<Element>().OfCategory( BuiltInCategorySets.PassPoints ).OfNotElementType().Where( filter ).OfType<FamilyInstance>() ) {
        if ( e.IsFamilyInstanceOf( RoutingFamilyType.PassPoint ) ) continue ;
        if ( e.TryGetProperty( RoutingParameter.RelatedPassPointId, out int id ) && id == passPointId ) yield return e ;
      }
    }

    public static string? GetRouteName( this Element element )
    {
      if ( false == element.TryGetProperty( RoutingParameter.RouteName, out string? value ) ) return null ;
      return value ;
    }

    public static int? GetSubRouteIndex( this Element element )
    {
      if ( false == element.TryGetProperty( RoutingParameter.SubRouteIndex, out int value ) ) return null ;
      return value ;
    }

    public static SubRouteInfo? GetSubRouteInfo( this Element element )
    {
      if ( element.GetRouteName() is not { } routeName ) return null ;
      if ( element.GetSubRouteIndex() is not { } subRouteIndex ) return null ;

      return new SubRouteInfo( routeName, subRouteIndex ) ;
    }

    public static IEnumerable<Connector> CollectRoutingEndPointConnectors( this Document document, string routeName, bool fromConnector )
    {
      return document.GetAllElementsOfRouteName<MEPCurve>( routeName ).SelectMany( e => e.GetRoutingConnectors( fromConnector ) ).Distinct() ;
    }

    public static (IReadOnlyCollection<Connector> From, IReadOnlyCollection<Connector>To) GetConnectors( this Document document, string routeName )
    {
      var fromList = document.CollectRoutingEndPointConnectors( routeName, true ).EnumerateAll() ;
      var toList = document.CollectRoutingEndPointConnectors( routeName, false ).EnumerateAll() ;
      return ( From: fromList, To: toList ) ;
    }

    public static IEnumerable<(MEPCurve, SubRoute)> CollectAllMultipliedRoutingElements( this Document document, int multiplicity )
    {
      if ( multiplicity < 2 ) throw new ArgumentOutOfRangeException( nameof( multiplicity ) ) ;

      var routes = RouteCache.Get( document ) ;

      foreach ( var mepCurve in document.GetAllElementsOfRoute<MEPCurve>() ) {
        if ( mepCurve.GetSubRouteInfo() is not { } subRouteInfo ) continue ;
        if ( mepCurve.GetRepresentativeSubRoute() != subRouteInfo ) continue ;
        if ( routes.GetSubRoute( subRouteInfo ) is not { } subRoute ) continue ;
        if ( subRoute.GetMultiplicity() < multiplicity ) continue ;

        yield return ( mepCurve, subRoute ) ;
      }
    }

    #endregion

    #region Routing (From-To)

    public static void SetRoutedElementFromToConnectorIds( this Element element, IReadOnlyCollection<int> fromIds, IReadOnlyCollection<int> toIds )
    {
      element.SetProperty( RoutingParameter.RoutedElementFromSideConnectorIds, string.Join( "|", fromIds ) ) ;
      element.SetProperty( RoutingParameter.RoutedElementToSideConnectorIds, string.Join( "|", toIds ) ) ;
    }

    public static IReadOnlyCollection<Connector> GetRoutingConnectors( this Element element, bool isFrom )
    {
      var manager = element.GetConnectorManager() ;
      if ( null == manager ) return Array.Empty<Connector>() ;

      var routingParam = ( isFrom ? RoutingParameter.RoutedElementFromSideConnectorIds : RoutingParameter.RoutedElementToSideConnectorIds ) ;
      if ( false == element.TryGetProperty( routingParam, out string? value ) ) return Array.Empty<Connector>() ;
      if ( null == value ) return Array.Empty<Connector>() ;

      var list = new List<Connector>() ;
      foreach ( var s in value.Split( '|' ) ) {
        if ( false == int.TryParse( s, out var id ) ) continue ;

        var conn = manager.Lookup( id ) ;
        if ( null == conn ) continue ;

        list.Add( conn ) ;
      }

      return list ;
    }

    public static bool IsRoutingConnector( this Connector connector, bool isFrom )
    {
      var routingParam = ( isFrom ? RoutingParameter.RoutedElementFromSideConnectorIds : RoutingParameter.RoutedElementToSideConnectorIds ) ;
      if ( false == connector.Owner.TryGetProperty( routingParam, out string? value ) ) return false ;
      if ( null == value ) return false ;

      var targetId = connector.Id ;
      return value.Split( '|' ).Any( s => int.TryParse( s, out var id ) && id == targetId ) ;
    }

    public static IEnumerable<Route> CollectRoutes( this Document document, AddInType addInType )
    {
      var routes = RouteCache.Get( document ).Values ;

      return addInType switch
      {
        AddInType.Electrical => routes.Where( r => r.GetSystemClassificationInfo().AddInType == AddInType.Electrical ),
        AddInType.Mechanical => routes.Where( r => r.GetSystemClassificationInfo().AddInType == AddInType.Mechanical ),
        AddInType.Undefined => routes.Where( r => r.GetSystemClassificationInfo().AddInType == AddInType.Undefined ),
        _ => routes.Where( r => r.GetSystemClassificationInfo().AddInType == AddInType.Undefined )
      } ;
    }

    public static IEnumerable<IEndPoint> GetNearestEndPoints( this Element element, bool isFrom )
    {
      if ( false == element.TryGetProperty( isFrom ? RoutingParameter.NearestFromSideEndPoints : RoutingParameter.NearestToSideEndPoints, out string? str ) ) {
        return Array.Empty<IEndPoint>() ;
      }

      if ( null == str ) {
        return Array.Empty<IEndPoint>() ;
      }

      return element.Document.ParseEndPoints( str ) ;
    }

    public static IReadOnlyCollection<SubRoute> GetSubRouteGroup( this Element element )
    {
      if ( ( element.GetRepresentativeSubRoute() ?? element.GetSubRouteInfo() ) is not { } subRouteInfo ) return Array.Empty<SubRoute>() ;

      var routeCache = RouteCache.Get( element.Document ) ;
      if ( routeCache.GetSubRoute( subRouteInfo ) is not { } subRoute ) return Array.Empty<SubRoute>() ;

      var subRouteGroup = subRoute.GetSubRouteGroup() ;
      if ( subRouteGroup.Count < 2 ) return new[] { subRoute } ;

      var result = new List<SubRoute>( subRouteGroup.Count ) ;
      result.AddRange( subRouteGroup.Select( routeCache.GetSubRoute ).NonNull() ) ;
      return result ;
    }

    public static void SetRepresentativeSubRoute( this Element element, SubRouteInfo subRouteInfo )
    {
      element.SetProperty( RoutingParameter.RepresentativeRouteName, subRouteInfo.RouteName ) ;
      element.SetProperty( RoutingParameter.RepresentativeSubRouteIndex, subRouteInfo.SubRouteIndex ) ;
    }

    public static string? GetRepresentativeRouteName( this Element element )
    {
      if ( false == element.TryGetProperty( RoutingParameter.RepresentativeRouteName, out string? value ) ) return null ;
      return value ;
    }

    public static SubRouteInfo? GetRepresentativeSubRoute( this Element element )
    {
      if ( element.GetRepresentativeRouteName() is not { } routeName ) return null ;
      if ( false == element.TryGetProperty( RoutingParameter.RepresentativeSubRouteIndex, out int subRouteIndex ) ) return null ;

      return new SubRouteInfo( routeName, subRouteIndex ) ;
    }

    /// <summary>
    /// 分岐管(Tee, Cross)のBranch側のRoute名を取得する
    /// </summary>
    /// <param name="elemnt"></param>
    /// <returns></returns>
    public static IEnumerable<string> GetBranchRouteNames( this Element element )
    {
      if ( false == element.TryGetProperty( RoutingParameter.BranchRouteNames, out string? str )) yield break;
      if ( str == null ) yield break;

      foreach ( var routeName in RouteNamesUtil.ParseRouteNames( str ) ) {
        yield return routeName ;
      }
    }
    
    #endregion

    #region Center Lines

    private static readonly ElementFilter CenterLineFilter = new ElementMulticategoryFilter( BuiltInCategorySets.CenterLineCategories ) ;

    public static IEnumerable<Element> GetCenterLine( this Element element )
    {
      var document = element.Document ;
      return element.GetDependentElements( CenterLineFilter ).Select( document.GetElement ).Where( e => e.IsValidObject ) ;
    }

    #endregion

    #region Shafts

    public static XYZ GetShaftPosition( this Opening shaft )
    {
      var box = shaft.get_BoundingBox( null ) ;
      return ( box.Min + box.Max ) * 0.5 ;
    }

    #endregion

    #region General

    private static FamilyInstance CreateFamilyInstance( this Document document, RoutingFamilyType familyType, XYZ position, StructuralType structuralType, bool useLevel, Level? level )
    {
      return document.CreateFamilyInstance( familyType, position, null, structuralType, useLevel, level ) ;
    }

    private static FamilyInstance CreateFamilyInstance( this Document document, RoutingFamilyType familyType, XYZ position, XYZ? direction, StructuralType structuralType, bool useLevel, Level? level )
    {
      var symbol = document.GetFamilySymbols( familyType ).FirstOrDefault() ?? throw new InvalidOperationException() ;
      if ( false == symbol.IsActive ) {
        symbol.Activate() ;
      }

      if ( false == useLevel ) {
        return document.Create.NewFamilyInstance( position, symbol, structuralType ) ;
      }

      level ??= document.GuessLevel( position ) ;
      var instance = document.Create.NewFamilyInstance( position, symbol, level, structuralType ) ;
      instance.get_Parameter( BuiltInParameter.INSTANCE_ELEVATION_PARAM ).Set( 0.0 ) ;

      if ( null != direction ) {
        var elevationAngle = Math.Atan2( direction.Z, Math.Sqrt( direction.X * direction.X + direction.Y * direction.Y ) ) ;
        var rotationAngle = Math.Atan2( direction.Y, direction.X ) ;

        ElementTransformUtils.RotateElement( document, instance.Id, Line.CreateBound( position, position + XYZ.BasisY ), -elevationAngle ) ;
        ElementTransformUtils.RotateElement( document, instance.Id, Line.CreateBound( position, position + XYZ.BasisZ ), rotationAngle ) ;
      }

      document.Regenerate() ;
      ElementTransformUtils.MoveElement( document, instance.Id, position - instance.GetTotalTransform().Origin ) ;

      return instance ;
    }

    private static readonly (BuiltInParameter, BuiltInParameter)[] LevelBuiltInParameterPairs = { ( BuiltInParameter.RBS_START_LEVEL_PARAM, BuiltInParameter.RBS_END_LEVEL_PARAM ) } ;

    public static ElementId GetLevelId( this Element element )
    {
      var levelId = element.LevelId ;
      if ( ElementId.InvalidElementId != levelId ) return levelId ;

      return LevelBuiltInParameterPairs.Select( tuple => GetUniqueParamLevelId( element, tuple.Item1, tuple.Item2 ) ).FirstOrDefault( paramLevelId => ElementId.InvalidElementId != paramLevelId ) ?? ElementId.InvalidElementId ;

      static ElementId GetUniqueParamLevelId( Element element, BuiltInParameter startParameter, BuiltInParameter endParameter )
      {
        var startLevelId = GetParamLevelId( element, startParameter ) ;
        var endLevelId = GetParamLevelId( element, endParameter ) ;
        if ( ElementId.InvalidElementId == startLevelId ) return ElementId.InvalidElementId ; // No levels
        if ( ElementId.InvalidElementId != endLevelId && startLevelId != endLevelId ) return ElementId.InvalidElementId ; // Different levels
        return startLevelId ;
      }

      static ElementId GetParamLevelId( Element element, BuiltInParameter builtInParameter )
      {
        if ( element.get_Parameter( builtInParameter ) is not { StorageType: StorageType.ElementId, HasValue: true } param ) return ElementId.InvalidElementId ;
        var elmId = param.AsElementId() ?? ElementId.InvalidElementId ;
        if ( ElementId.InvalidElementId == elmId ) return ElementId.InvalidElementId ;

        return element.Document.GetElementById<Level>( elmId )?.Id ?? ElementId.InvalidElementId ;
      }
    }

    public static ElementId GuessLevelId( this Document document, XYZ position )
    {
      return GuessLevel( document, position ).Id ;
    }

    public static Level GuessLevel( this Document document, XYZ position )
    {
      var z = position.Z - document.Application.VertexTolerance ;
      var list = document.GetAllElements<Level>().Select( level => new LevelByElevation( level.Elevation, level ) ).ToList() ;
      if ( 0 == list.Count ) Level.Create( document, 0 ) ;

      list.Sort() ;

      var index = list.BinarySearch( new LevelByElevation( z, null ) ) ;
      if ( 0 <= index ) return list[ index ].Level! ;

      var greaterIndex = ~index ;
      return list[ Math.Max( 0, greaterIndex - 1 ) ].Level! ;
    }

    private record LevelByElevation( double LevelElevation, Level? Level ) : IComparable<LevelByElevation>, IComparable
    {
      public int CompareTo( LevelByElevation? other )
      {
        if ( ReferenceEquals( this, other ) ) return 0 ;
        if ( ReferenceEquals( null, other ) ) return 1 ;
        return LevelElevation.CompareTo( other.LevelElevation ) ;
      }

      public int CompareTo( object? obj )
      {
        if ( ReferenceEquals( null, obj ) ) return 1 ;
        if ( ReferenceEquals( this, obj ) ) return 0 ;

        return obj is LevelByElevation other ? CompareTo( other ) : throw new ArgumentException( $"Object must be of type {nameof( LevelByElevation )}" ) ;
      }
    }

    #endregion
  }
}