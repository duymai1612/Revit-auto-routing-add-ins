using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.EndPoints ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Connector information which is not be invalidated by rolling back.
  /// </summary>
  public readonly struct ConnectorId : IEquatable<ConnectorId>
  {
    private int OwnerIdValue { get ; }
    private int Id { get ; }

    public ConnectorId( Connector connector )
    {
      OwnerIdValue = connector.Owner.Id.IntegerValue ;
      Id = connector.Id ;
    }

    public ConnectorId( ConnectorEndPoint connectorEndPoint )
    {
      OwnerIdValue = connectorEndPoint.EquipmentId.IntegerValue ;
      Id = connectorEndPoint.ConnectorIndex ;
    }

    public Element? GetOwner( Document document )
    {
      return document.GetElement( new ElementId( OwnerIdValue ) ) ;
    }
    public Connector? GetConnector( Document document )
    {
      return GetOwner( document )?.GetConnectorManager()?.Lookup( Id ) ;
    }
    public Connector? GetConnector( Element element )
    {
      if ( element.Id.IntegerValue != OwnerIdValue ) return null ;

      return element.GetConnectorManager()?.Lookup( Id ) ;
    }

    public static IReadOnlyList<Connector> ToConnectorList( Document document, IReadOnlyCollection<ConnectorId> collection )
    {
      var list = new List<Connector>( collection.Count ) ;
      foreach ( var connId in collection ) {
        if ( connId.GetConnector( document ) is not { } connector ) continue ;

        list.Add( connector ) ;
      }

      return list ;
    }

    public bool Equals( ConnectorId other )
    {
      return OwnerIdValue.Equals( other.OwnerIdValue ) && Id == other.Id ;
    }

    public override bool Equals( object? obj )
    {
      if ( obj is not ConnectorId another ) return false ;

      return Equals( another ) ;
    }

    public override int GetHashCode()
    {
      unchecked {
        return ( OwnerIdValue.GetHashCode() * 397 ) ^ Id ;
      }
    }

    public static bool operator ==( ConnectorId? left, ConnectorId? right )
    {
      return Equals( left, right ) ;
    }

    public static bool operator !=( ConnectorId? left, ConnectorId? right )
    {
      return ! Equals( left, right ) ;
    }
  }
}