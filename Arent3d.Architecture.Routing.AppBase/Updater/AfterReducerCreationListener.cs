using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Updater
{
  public class AfterReducerCreationListener : IDocumentUpdateListener
  {
    public Guid Guid { get ; } = new Guid( "{C1339C96-074F-473F-82FE-D050943AEDFA}" ) ;
    public string Name => nameof( AfterReducerCreationListener ) ;
    public string Description => nameof( AfterReducerCreationListener ) ;

    public ChangePriority ChangePriority => ChangePriority.MEPSystems ;
    public DocumentUpdateListenType ListenType => DocumentUpdateListenType.Addition ;

    public ElementFilter GetElementFilter()
    {
      return new ElementMulticategoryFilter( BuiltInCategorySets.Fittings ) ;
    }

    public IEnumerable<ParameterProxy> GetListeningParameters() => throw new NotSupportedException() ;

    public void Execute( UpdaterData data )
    {
      var document = data.GetDocument() ;
      foreach ( var elementId in data.GetAddedElementIds() ) {
        if ( document.GetElementById<FamilyInstance>( elementId ) is not { } fitting ) continue ;
        if ( null != fitting.GetRouteName() ) continue ;  // already added
        if ( GetNeighborRoutingParameters( fitting ) is not { } parameters ) continue ;

        parameters.ApplyTo( fitting ) ;
      }
    }

    private static RoutingParameters? GetNeighborRoutingParameters( FamilyInstance fitting )
    {
      foreach ( var fittingConnector in fitting.GetConnectors().OfEnd() ) {
        foreach ( var anotherConnector in fittingConnector.GetConnectedConnectors().OfEnd() ) {
          var elm = anotherConnector.Owner ;
          if ( false == elm.IsValidObject ) continue ;

          if ( RoutingParameters.Create( elm ) is not { } parameters ) continue ;
          if ( parameters.EndsWith( fittingConnector ) ) continue ;

          return parameters ;
        }
      }

      return null ;
    }

    private class RoutingParameters
    {
      public static RoutingParameters? Create( Element elm )
      {
        if ( elm.GetRouteName() is not { } routeName ) return null ;
        if ( elm.GetSubRouteIndex() is not { } subRouteIndex ) return null ;
        if ( false == elm.TryGetProperty( RoutingParameter.NearestFromSideEndPoints, out string? fromSide ) ) return null ;
        if ( false == elm.TryGetProperty( RoutingParameter.NearestToSideEndPoints, out string? toSide ) ) return null ;

        return new RoutingParameters( routeName, subRouteIndex, fromSide!, toSide! ) ;
      }

      private string RouteName { get ; }
      private int SubRouteIndex { get ; }
      private string RoutedElementFromSideConnectorIds { get ; }
      private string RoutedElementToSideConnectorIds { get ; }

      private RoutingParameters(  string routeName, int subRouteIndex, string fromSide, string toSide )
      {
        RouteName = routeName ;
        SubRouteIndex = subRouteIndex ;
        RoutedElementFromSideConnectorIds = fromSide ;
        RoutedElementToSideConnectorIds = toSide ;
      }

      public bool EndsWith( Connector nextConnector )
      {
        var endPoints1 = nextConnector.Owner.Document.ParseEndPoints( RoutedElementFromSideConnectorIds ) ;
        var endPoints2 = nextConnector.Owner.Document.ParseEndPoints( RoutedElementToSideConnectorIds ) ;
        return endPoints1.Concat( endPoints2 ).Any( ep => HasEndConnector( ep, nextConnector ) ) ;
      }

      private static bool HasEndConnector( IEndPoint ep, Connector nextConnector )
      {
        return ep switch
        {
          ConnectorEndPoint c => c.EquipmentId == nextConnector.Owner.Id && c.ConnectorIndex == nextConnector.Id,
          PassPointBranchEndPoint => false,
          PassPointEndPoint => false, // FIXME: check whether >|= type pass point or =|< type pass point
          RouteEndPoint r => r.RouteName == nextConnector.Owner.GetRouteName() && r.SubRouteIndex == nextConnector.Owner.GetSubRouteIndex(),
          TerminatePointEndPoint t => false,
          _ => false,
        } ;
      }

      public void ApplyTo( Element reducer )
      {
        reducer.SetProperty( RoutingParameter.RouteName, RouteName ) ;
        reducer.SetProperty( RoutingParameter.SubRouteIndex, SubRouteIndex ) ;
        reducer.SetProperty( RoutingParameter.NearestFromSideEndPoints, RoutedElementFromSideConnectorIds ) ;
        reducer.SetProperty( RoutingParameter.NearestToSideEndPoints, RoutedElementToSideConnectorIds ) ;
      }
    }
  }
}