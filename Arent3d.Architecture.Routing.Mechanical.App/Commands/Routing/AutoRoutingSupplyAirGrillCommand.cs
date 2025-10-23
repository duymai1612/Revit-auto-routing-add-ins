using System.Collections.Generic ;
using System.Linq ;
using System.Text.RegularExpressions ;
using Arent3d.Architecture.Routing.AppBase ;
using Arent3d.Architecture.Routing.AppBase.Commands.Routing ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit.UI ;
using Autodesk.Revit.Attributes ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;
using Autodesk.Revit.UI.Selection ;

namespace Arent3d.Architecture.Routing.Mechanical.App.Commands.Routing
{
  [Transaction( TransactionMode.Manual )]
  [DisplayNameKey( "Mechanical.App.Commands.Routing.AutoRoutingSupplyAirGrillCommand",
    DefaultString = "Auto Routing\nSupply Air Grill" )]
  [Image( "resources/AutoRoutingSupplyAirGrill.png" )]
  public class AutoRoutingSupplyAirGrillCommand : RoutingCommandBase<AutoRoutingSupplyAirGrillCommand.AutoRoutingState>
  {
    public class AutoRoutingState
    {
      public List<(Connector OutConnector, Connector SupplyAirGrillConnector)> Pairings { get ; }

      public AutoRoutingState( List<(Connector, Connector)> pairings )
      {
        Pairings = pairings ;
      }
    }

    protected override string GetTransactionNameKey() => "TransactionName.Commands.Routing.AutoRoutingSupplyAirGrill" ;

    protected override RoutingExecutor CreateRoutingExecutor( Document document, View view ) =>
      AppCommandSettings.CreateRoutingExecutor( document, view ) ;

    protected override OperationResult<AutoRoutingState> OperateUI( ExternalCommandData commandData,
      ElementSet elements )
    {
      var uiDocument = commandData.Application.ActiveUIDocument ;
      var document = uiDocument.Document ;

      try {
        // Step 1: Pick AIR DISTRIBUTION BOX
        var reference = uiDocument.Selection.PickObject( ObjectType.Element,
          new AutoRoutingSupplyAirGrillCommandUtil.AirDistributionBoxSelectionFilter(),
          "Select AIR DISTRIBUTION BOX" ) ;

        var airDistributionBox = document.GetElement( reference ) as FamilyInstance ;
        if ( null == airDistributionBox ) return OperationResult<AutoRoutingState>.Cancelled ;

        // Step 2: Get In Connector and Out Connectors
        var (inConnector, outConnectors) =
          AutoRoutingSupplyAirGrillCommandUtil.GetAirDistributionBoxConnectors( airDistributionBox ) ;
        if ( null == inConnector || outConnectors.Count == 0 ) {
          TaskDialog.Show( "Error", "AIR DISTRIBUTION BOX must have In Connector and Out Connectors." ) ;
          return OperationResult<AutoRoutingState>.Cancelled ;
        }

        // Step 3: Find all Supply Air Grills with same Duct System
        var supplyAirGrills = AutoRoutingSupplyAirGrillCommandUtil.FindSupplyAirGrills( document, inConnector ) ;
        if ( supplyAirGrills.Count == 0 ) {
          TaskDialog.Show( "Info", "No Supply Air Grills found with matching Duct System." ) ;
          return OperationResult<AutoRoutingState>.Cancelled ;
        }

        // Step 4: Create pairings
        var pairings =
          AutoRoutingSupplyAirGrillCommandUtil.CreatePairings( inConnector, outConnectors, supplyAirGrills ) ;

        return new OperationResult<AutoRoutingState>( new AutoRoutingState( pairings ) ) ;
      }
      catch ( Autodesk.Revit.Exceptions.OperationCanceledException ) {
        return OperationResult<AutoRoutingState>.Cancelled ;
      }
    }

    protected override IReadOnlyCollection<(string RouteName, RouteSegment Segment)> GetRouteSegments(
      Document document, AutoRoutingState state )
    {
      var segments = new List<(string RouteName, RouteSegment Segment)>() ;
      var routes = RouteCache.Get( document ) ;

      foreach ( var (outConnector, grillConnector) in state.Pairings ) {
        var classificationInfo = MEPSystemClassificationInfo.From( outConnector ) ;
        if ( null == classificationInfo ) continue ;

        var systemType = RouteMEPSystem.GetSystemType( document, outConnector ) ;
        var curveType = RouteMEPSystem.GetMEPCurveType( document, new[] { outConnector }, systemType ) ;
        var diameter = outConnector.GetDiameter() ;

        var fromEndPoint = new ConnectorEndPoint( outConnector, null ) ;
        var toEndPoint = new ConnectorEndPoint( grillConnector, null ) ;

        var nameBase = systemType?.Name ?? curveType.Category.Name ;
        var nextIndex = GetRouteNameIndex( routes, nameBase ) ;
        var routeName = nameBase + "_" + nextIndex ;
        routes.FindOrCreate( routeName ) ;

        var segment = new RouteSegment( classificationInfo, systemType, curveType, fromEndPoint, toEndPoint, diameter,
          isRoutingOnPipeSpace: false, fromFixedHeight: null, toFixedHeight: null, avoidType: AvoidType.Whichever,
          shaftElementId: ElementId.InvalidElementId ) ;

        segments.Add( ( routeName, segment ) ) ;
      }

      return segments ;
    }

    private static int GetRouteNameIndex( RouteCache routes, string? targetName )
    {
      string pattern = @"^" + Regex.Escape( targetName ?? string.Empty ) + @"_(\d+)$" ;
      var regex = new Regex( pattern ) ;

      var lastIndex = routes.Keys.Select( k => regex.Match( k ) ).Where( m => m.Success )
        .Select( m => int.Parse( m.Groups[ 1 ].Value ) ).Append( 0 ).Max() ;

      return lastIndex + 1 ;
    }

    protected override void AfterRouteGenerated( Document document, IReadOnlyCollection<Route> executeResultValue )
    {
      // Additional processing after routes are generated
    }
  }
}
