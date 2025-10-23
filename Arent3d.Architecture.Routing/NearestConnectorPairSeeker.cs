using System.Collections.Generic ;
using System.Linq ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public class NearestConnectorPairSeeker
  {
    private readonly List<Connector> _connectors ;
    private readonly List<ConnectorElement> _connectorElements ;
    private readonly List<DistanceInfo> _distances = new() ;

    public NearestConnectorPairSeeker( IEnumerable<Connector> connectors, IEnumerable<ConnectorElement> connectorElements, Transform familyInstanceTransform )
    {
      _connectors = connectors.ToList() ;
      _connectorElements = connectorElements.ToList() ;

      foreach ( var connectorElement in _connectorElements ) {
        var domain = connectorElement.Domain ;
        var systemClassification = connectorElement.SystemClassification ;
        var origin = familyInstanceTransform.OfPoint( connectorElement.Origin ) ;
        var dirZ = familyInstanceTransform.OfVector( connectorElement.CoordinateSystem.BasisZ ) ;
        var dirX = familyInstanceTransform.OfVector( connectorElement.CoordinateSystem.BasisX ) ;
        foreach ( var connector in _connectors.Where( c => domain == c.Domain ) ) {
          if ( domain != Domain.DomainCableTrayConduit && false == connector.HasCompatibleSystemType( systemClassification ) ) continue ;

          var distance = connector.Origin.DistanceTo( origin ) ;
          var angleZ = connector.CoordinateSystem.BasisZ.AngleTo( dirZ ) ;
          var angleX = connector.CoordinateSystem.BasisX.AngleTo( dirX ) ;
          _distances.Add( new DistanceInfo( connector, connectorElement, distance, angleZ, angleX ) ) ;
        }
      }

      _distances.Sort( DistanceInfo.Compare ) ;
    }

    public (Connector?, ConnectorElement?) Pop()
    {
      if ( 0 == _distances.Count ) return ( null, null ) ;

      var first = _distances[ 0 ] ;
      var (conn, connElm) = ( first.Connector, first.ConnectorElement ) ;
      _distances.RemoveAll( d => d.IsConnector( conn ) || d.IsConnectorElement( connElm ) ) ;

      return ( conn, connElm ) ;
    }

    private class DistanceInfo
    {
      public Connector Connector { get ; }
      public ConnectorElement ConnectorElement { get ; }
      private double Distance { get ; }
      private double DirectionalDistance { get ; }

      public DistanceInfo( Connector connector, ConnectorElement connectorElement, double distance, double angleZ, double angleX )
      {
        Connector = connector ;
        ConnectorElement = connectorElement ;
        Distance = distance ;
        DirectionalDistance = angleZ + angleX ;
      }

      public bool IsConnector( Connector conn )
      {
        return ( conn.Owner.Id == Connector.Owner.Id && conn.Id == Connector.Id ) ;
      }

      public bool IsConnectorElement( ConnectorElement connElm )
      {
        return ( connElm.Id == ConnectorElement.Id ) ;
      }

      public static int Compare( DistanceInfo x, DistanceInfo y )
      {
        var dist = x.Distance.CompareTo( y.Distance ) ;
        if ( 0 != dist ) return dist ;

        var dir = x.DirectionalDistance.CompareTo( y.DirectionalDistance ) ;
        if ( 0 != dir ) return dir ;

        var elm = x.ConnectorElement.Id.IntegerValue.CompareTo( y.ConnectorElement.Id.IntegerValue ) ;
        if ( 0 != elm ) return elm ;

        var connElm = x.Connector.Owner.Id.IntegerValue.CompareTo( y.Connector.Owner.Id.IntegerValue ) ;
        if ( 0 != connElm ) return connElm ;

        var conn = x.Connector.Id.CompareTo( y.Connector.Id ) ;
        if ( 0 != conn ) return conn ;

        return 0 ;
      }
    }
  }
}