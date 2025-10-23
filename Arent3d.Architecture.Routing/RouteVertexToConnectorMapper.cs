using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Routing ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Mapper class of <see cref="IRouteVertex"/> to <see cref="Connector"/>.
  /// </summary>
  public class RouteVertexToConnectorMapper
  {
    private readonly Dictionary<IRouteVertex, List<ConnectorId>> _dicMap = new() ;

    /// <summary>
    /// Adds a connector for an auto routing vertex.
    /// </summary>
    /// <param name="routeVertex">An auto routing vertex.</param>
    /// <param name="connector">A connector.</param>
    public void Add( IRouteVertex routeVertex, Connector connector )
    {
      if ( false == _dicMap.TryGetValue( routeVertex, out var connectors ) ) {
        connectors = new List<ConnectorId>() ;
        _dicMap.Add( routeVertex, connectors ) ;
      }

      connectors.Add( new ConnectorId( connector ) ) ;
    }

    /// <summary>
    /// Returns a position of a connector who will be newly created.
    /// </summary>
    /// <param name="routeVertex">An auto routing vertex where a connector is to be created.</param>
    /// <param name="anotherRouteVertex">An opposite side of auto routing vertex where a connector is to be created.</param>
    /// <returns></returns>
    public Vector3d GetNewConnectorPosition( IRouteVertex routeVertex, IRouteVertex anotherRouteVertex )
    {
      return routeVertex.Position ;
    }

    private static Vector3d GetOffset( Vector3d anotherEndDir )
    {
      // offset can be a vector whose norm is 1.
      // c.f.) Revit SDK Sample `AutoRoute/Command.cs', min1FittingLength
      return anotherEndDir.normalized ;
    }

    public IEnumerable<IReadOnlyList<Connector>> GetConnections( Document document )
    {
      return _dicMap.Values.Select( list => ConnectorId.ToConnectorList( document, list ) ) ;
    }
  }
}