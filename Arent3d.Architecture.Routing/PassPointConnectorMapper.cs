using System ;
using System.Collections.Generic ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Manages relations between pass points and connectors.
  /// </summary>
  public class PassPointConnectorMapper
  {
    private readonly Dictionary<(Route Route, int PassPointElementId), (ConnectorId? Prev, ConnectorId? Next)> _passPointConnectors = new() ;

    public void Add( Route route, ElementId passPointElementId, Connector connector, bool continuesToFrom )
    {
      var key = ( Route: route, PassPointElementId: passPointElementId.IntegerValue ) ;

      if ( _passPointConnectors.TryGetValue( key, out var tuple ) ) {
        if ( continuesToFrom ) {
          if ( null != tuple.Next ) throw new InvalidOperationException() ;
          _passPointConnectors[ key ] = ( tuple.Prev, new ConnectorId( connector ) ) ;
        }
        else {
          if ( null != tuple.Prev ) throw new InvalidOperationException() ;
          _passPointConnectors[ key ] = ( new ConnectorId( connector ), tuple.Next ) ;
        }
      }
      else {
        if ( continuesToFrom ) {
          _passPointConnectors.Add( key, ( null, new ConnectorId( connector ) ) ) ;
        }
        else {
          _passPointConnectors.Add( key, ( new ConnectorId( connector ), null ) ) ;
        }
      }
    }

    public void Merge( PassPointConnectorMapper another )
    {
      foreach ( var (key, (prevConnector, nextConnector)) in another._passPointConnectors ) {
        if ( _passPointConnectors.TryGetValue( key, out var tuple ) ) {
          if ( null != tuple.Prev && null != prevConnector ) throw new InvalidOperationException() ;
          if ( null != tuple.Next && null != nextConnector ) throw new InvalidOperationException() ;

          _passPointConnectors[ key ] = ( tuple.Prev ?? prevConnector, tuple.Next ?? nextConnector ) ;
        }
        else {
          _passPointConnectors.Add( key, ( prevConnector, nextConnector ) ) ;
        }
      }
    }

    public IEnumerable<(Route Route, ElementId PassPointElementId, Connector Prev, Connector Next)> GetPassPointConnections( Document document )
    {
      foreach ( var (key, (prevConnector, nextConnector)) in _passPointConnectors ) {
        if ( prevConnector?.GetConnector( document ) is not { } con1 ) continue ;
        if ( nextConnector?.GetConnector( document ) is not { } con2 ) continue ;

        yield return ( key.Route, new ElementId( key.PassPointElementId ), con1, con2 ) ;
      }
    }
  }
}