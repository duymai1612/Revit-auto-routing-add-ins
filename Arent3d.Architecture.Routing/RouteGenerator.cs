using System ;
using System.Collections.Generic ;
using System.IO ;
using System.Linq ;
using Arent3d.Architecture.Routing.CollisionTree ;
using Arent3d.Architecture.Routing.Extensions ;
using Arent3d.Architecture.Routing.FittingSizeCalculators ;
using Arent3d.Architecture.Routing.Storable ;
using Arent3d.Architecture.Routing.Storable.Model ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Routing ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Route generator class. This calculates route paths from routing targets and transforms revit elements.
  /// </summary>
  public class RouteGenerator : RouteGeneratorBase<AutoRoutingTarget>
  {
    private readonly Document _document ;
    private readonly IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> _routeConditions ;
    private readonly List<Connector[]> _badConnectors = new() ;
    private readonly PassPointConnectorMapper _globalPassPointConnectorMapper = new() ;

    public RouteGenerator( Document document, IReadOnlyCollection<Route> routes, AutoRoutingTargetGenerator autoRoutingTargetGenerator, IFittingSizeCalculator fittingSizeCalculator, ICollisionCheckTargetCollector collector )
    {
      _document = document ;

      _routeConditions = CreateRouteConditions( document, routes, fittingSizeCalculator ) ;
      var targets = autoRoutingTargetGenerator.Create( routes, _routeConditions ) ;
      RoutingTargetGroups = targets.ToList() ;
      ErasePreviousRoutes() ; // Delete before CollisionCheckTree is built.

      CollisionCheckTree = new CollisionTree.CollisionTree( document, collector, _routeConditions ) ;
      StructureGraph = DocumentMapper.Get( document ).RackCollection ;

      Specifications.Set( DiameterProvider.Instance, PipeClearanceProvider.Instance ) ;
    }

    private static IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> CreateRouteConditions( Document document, IReadOnlyCollection<Route> routes, IFittingSizeCalculator fittingSizeCalculator )
    {
      var dic = new Dictionary<SubRouteInfo, MEPSystemRouteCondition>() ;
      
      foreach ( var route in routes ) {
        foreach ( var subRoute in route.SubRoutes ) {
          var key = new SubRouteInfo( subRoute ) ;
          if ( dic.ContainsKey( key ) ) break ;  // same sub route

          var mepSystem = new RouteMEPSystem( document, subRoute ) ;

          var edgeDiameter = subRoute.GetDiameter() ;
          var spec = new MEPSystemPipeSpec( mepSystem, fittingSizeCalculator ) ;
          var routeCondition = new MEPSystemRouteCondition( spec, edgeDiameter, subRoute.AvoidType ) ;

          dic.Add( key, routeCondition ) ;
        }
      }

      return dic ;
    }

    public IReadOnlyCollection<Connector[]> GetBadConnectorSet() => _badConnectors ;

    protected override IReadOnlyList<IReadOnlyCollection<AutoRoutingTarget>> RoutingTargetGroups { get ; }

    protected override CollisionTree.CollisionTree CollisionCheckTree { get ; }

    protected override IStructureGraph StructureGraph { get ; }

    /// <summary>
    /// Erase all previous ducts and pipes in between routing targets.
    /// </summary>
    private void ErasePreviousRoutes()
    {
      EraseRoutes( _document, RoutingTargetGroups.SelectMany( group => group ).SelectMany( t => t.Routes ).Select( route => route.RouteName ), false ) ;
    }

    public static void EraseRoutes( Document document, IEnumerable<string> routeNames, bool eraseRouteStoragesAndPassPoints )
    {
      var hashSet = ( routeNames as ISet<string> ) ?? routeNames.ToHashSet() ;

      var list = document.GetAllElementsOfRoute<Element>().Where( e => e.GetRouteName() is { } routeName && hashSet.Contains( routeName ) ) ;
      if ( false == eraseRouteStoragesAndPassPoints ) {
        // do not erase pass points
        list = list.Where( p => false == ( p is FamilyInstance fi && ( fi.IsFamilyInstanceOf( RoutingFamilyType.PassPoint ) || fi.IsFamilyInstanceOf( RoutingFamilyType.TerminatePoint )) ) );
      }

      List<string> elementIds = list.Select( elm => elm.Id.IntegerValue.ToString() ).Distinct().ToList() ;
      RemoveRouteDetailSymbol( document, elementIds ) ;

      document.Delete( list.Select( elm => elm.Id ).Distinct().ToArray() ) ;

      if ( eraseRouteStoragesAndPassPoints ) {
        // erase routes, too.
        RouteCache.Get( document ).Drop( hashSet ) ;
      }
    }

    protected override void OnGenerationStarted()
    {
#if DUMP_LOGS
      RoutingTargets.DumpRoutingTargets( GetTargetsLogFileName( _document ), CollisionCheckTree ) ;
#endif

      // TODO, if needed
    }

    protected override void OnRoutingTargetProcessed( IReadOnlyCollection<AutoRoutingTarget> routingTargets, MergedAutoRoutingResult result )
    {
#if DUMP_LOGS
      result.DebugExport( GetResultLogFileName( _document, routingTarget ) ) ;
#endif

      var mepSystemCreator = new MEPSystemCreator( _document, routingTargets, _routeConditions ) ;

      foreach ( var routeVertex in result.RouteVertices ) {
        if ( routeVertex is not TerminalPoint ) continue ;

        mepSystemCreator.RegisterEndPointConnector( routeVertex ) ;
      }

      var newElements = CreateEdges( mepSystemCreator, result ).ToList() ;
      newElements.AddRange( mepSystemCreator.ConnectAllVertices() ) ;

      _document.Regenerate() ;

      _globalPassPointConnectorMapper.Merge( mepSystemCreator.PassPointConnectorMapper ) ;

      RegisterBadConnectors( mepSystemCreator.GetBadConnectorSet() ) ;
    }

    protected virtual IEnumerable<Element> CreateEdges( MEPSystemCreator mepSystemCreator, MergedAutoRoutingResult result )
    {
      return result.RouteEdges.Select( routeEdge => mepSystemCreator.CreateEdgeElement( routeEdge, mepSystemCreator.GetSubRoute( routeEdge ), result.GetPassingEndPointInfo( routeEdge ) ) ) ;
    }

    private void RegisterBadConnectors( IEnumerable<Connector[]> badConnectorSet )
    {
      _badConnectors.AddRange( badConnectorSet ) ;
    }

    protected override void OnGenerationFinished()
    {
      foreach ( var (route, passPointId, prevConn, nextConn) in _globalPassPointConnectorMapper.GetPassPointConnections( _document ) ) {
        var element = _document.GetElement( passPointId ) ;
        if ( element.GetRouteName() == route.RouteName ) {
          element.SetPassPointConnectors( new[] { prevConn }, new[] { nextConn } ) ;
        }

        var (success, fitting) = MEPSystemCreator.ConnectConnectors( _document, new[] { prevConn, nextConn } ) ;
        if ( success && null != fitting ) {
          // set routing id.
          fitting.SetProperty( RoutingParameter.RouteName, route.RouteName ) ;
          if ( ( prevConn.Owner.GetSubRouteIndex() ?? nextConn.Owner.GetSubRouteIndex() ) is { } subRouteIndex ) {
            fitting.SetProperty( RoutingParameter.SubRouteIndex, subRouteIndex ) ;
          }

          // Relate fitting to the pass point.
          element.SetProperty( RoutingParameter.RelatedPassPointId, passPointId.IntegerValue ) ;
        }
      }
    }

    private static string GetLogDirectoryName( Document document )
    {
      var dir = Path.Combine( Path.GetDirectoryName( document.PathName )!, Path.GetFileNameWithoutExtension( document.PathName ) ) ;
      return Directory.CreateDirectory( dir ).FullName ;
    }

    private static string GetTargetsLogFileName( Document document )
    {
      return Path.Combine( GetLogDirectoryName( document ), "RoutingTargets.xml" ) ;
    }

    private static string GetResultLogFileName( Document document, AutoRoutingTarget routingTarget )
    {
      return Path.Combine( GetLogDirectoryName( document ), routingTarget.LineId + ".log" ) ;
    }

    public static void CorrectEnvelopes( Document document )
    {
      // get all envelope
      var envelopes = document.GetAllFamilyInstances( RoutingFamilyType.Envelope ).ToList() ;
      if ( ! envelopes.Any() ) return ;
      var parentEnvelopes = envelopes.Where( f => string.IsNullOrEmpty( f.ParametersMap.get_Item( "Revit.Property.Builtin.ParentEnvelopeId".GetDocumentStringByKeyOrDefault( document, "Parent Envelope Id" ) ).AsString() ) ).ToList() ;
      if ( ! parentEnvelopes.Any() ) return ;
      
      // get offset value
      OffsetSettingStorable settingStorable = document.GetOffsetSettingStorable() ;
      var offset = settingStorable.OffsetSettingsData.Offset.MillimetersToRevitUnits() ;
      foreach ( var parentEnvelope in parentEnvelopes ) {
        var parentLocation = parentEnvelope.Location as LocationPoint ;
        var childrenEnvelope = envelopes.FirstOrDefault( f => f.ParametersMap.get_Item( "Revit.Property.Builtin.ParentEnvelopeId".GetDocumentStringByKeyOrDefault( document, "Parent Envelope Id" ) ).AsString() == parentEnvelope.Id.ToString() ) ;
        if ( childrenEnvelope == null ) continue ;
        var childrenLocation = childrenEnvelope.Location as LocationPoint ;
        childrenLocation!.Point = new XYZ( parentLocation!.Point.X, parentLocation!.Point.Y, parentLocation!.Point.Z - offset ) ;
      }
    }

    private static void RemoveRouteDetailSymbol( Document document, List<string> elementIds )
    {
      var detailSymbolStorable = document.GetAllStorables<DetailSymbolStorable>().FirstOrDefault() ?? document.GetDetailSymbolStorable() ;
      if ( ! detailSymbolStorable.DetailSymbolModelData.Any() ) return ;
      var detailSymbolModels = new List<DetailSymbolModel>() ;
      foreach ( var detailSymbolModel in detailSymbolStorable.DetailSymbolModelData.Where( d => elementIds.Contains( d.ConduitId ) ).ToList() ) {
        // delete symbol
        var symbolId = document.GetAllElements<Element>().OfCategory( BuiltInCategory.OST_TextNotes ).Where( e => e.Id.IntegerValue.ToString() == detailSymbolModel.DetailSymbolId ).Select( t => t.Id ).FirstOrDefault() ;
        if ( symbolId != null ) document.Delete( symbolId ) ;
        foreach ( var lineId in detailSymbolModel.LineIds.Split( ',' ) ) {
          var id = document.GetAllElements<Element>().OfCategory( BuiltInCategory.OST_Lines ).Where( e => e.Id.IntegerValue.ToString() == lineId ).Select( e => e.Id ).FirstOrDefault() ;
          if ( id != null ) document.Delete( id ) ;
        }

        detailSymbolModels.Add( detailSymbolModel ) ;
      }

      if ( ! detailSymbolModels.Any() ) return ;
      foreach ( var detailSymbolModel in detailSymbolModels ) {
        detailSymbolStorable.DetailSymbolModelData.Remove( detailSymbolModel ) ;
      }

      var detailSymbols = detailSymbolModels.Select( d => d.DetailSymbolId ).Distinct().ToList() ;
      if ( detailSymbolStorable.DetailSymbolModelData.Any() && detailSymbols.Count == 1 ) {
        var detailSymbolModel = detailSymbolModels.FirstOrDefault() ;
        if ( detailSymbolModel!.IsParentSymbol ) {
          var detailSymbolModelParent = detailSymbolStorable.DetailSymbolModelData.FirstOrDefault( d => d.DetailSymbol == detailSymbolModel.DetailSymbol && d.Code == detailSymbolModel.Code && d.IsParentSymbol ) ;
          if ( detailSymbolModelParent == null ) {
            UpdateSymbolOfConduitSameSymbolAndDifferentCode( document, detailSymbolStorable.DetailSymbolModelData, detailSymbolModel.DetailSymbol, detailSymbolModel.Code ) ;
          }
        }
      }

      detailSymbolStorable.Save() ;
    }

    private static void UpdateSymbolOfConduitSameSymbolAndDifferentCode( Document doc, List<DetailSymbolModel> detailSymbolModels, string detailSymbol, string code )
    {
      var firstChildSymbol = detailSymbolModels.FirstOrDefault( d => d.DetailSymbol == detailSymbol && d.Code != code ) ;
      if ( firstChildSymbol == null ) return ;
      {
        var detailSymbolIds = detailSymbolModels.Where( d => d.DetailSymbol == firstChildSymbol.DetailSymbol && d.Code == firstChildSymbol.Code ).Select( d => d.DetailSymbolId ).Distinct().ToList() ;
        foreach ( var id in detailSymbolIds ) {
          var textElement = doc.GetAllElements<Element>().OfCategory( BuiltInCategory.OST_TextNotes ).FirstOrDefault( t => t.Id.IntegerValue.ToString() == id ) ;
          if ( textElement == null ) continue ;
          var textNote = ( textElement as TextNote ) ! ;
          CreateNewTextNoteType( doc, textNote, 0 ) ;
        }

        foreach ( var detailSymbolModel in detailSymbolModels.Where( d => d.DetailSymbol == firstChildSymbol.DetailSymbol && d.Code == firstChildSymbol.Code ).ToList() ) {
          detailSymbolModel.IsParentSymbol = true ;
        }
      }
    }

    private static void CreateNewTextNoteType( Document doc, TextNote textNote, int color )
    {
      //Create new text type
      string strStyleName = textNote.TextNoteType.Name + "-" + color ;

      var textNoteType = new FilteredElementCollector( doc ).OfClass( typeof( TextNoteType ) ).WhereElementIsElementType().Cast<TextNoteType>().FirstOrDefault( tt => Equals( strStyleName, tt.Name ) ) ;
      if ( textNoteType == null ) {
        // Create new Note type
        Element ele = textNote.TextNoteType.Duplicate( strStyleName ) ;
        textNoteType = ( ele as TextNoteType ) ! ;
        textNoteType.get_Parameter( BuiltInParameter.LINE_COLOR ).Set( color ) ;
      }

      // Change the text notes type to the new type
      textNote.ChangeTypeId( textNoteType!.Id ) ;
    }
  }
}