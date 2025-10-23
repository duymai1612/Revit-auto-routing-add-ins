using System.Diagnostics ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.EndPoints
{
  [DebuggerDisplay("{Key}")]
  public class ConnectorEndPoint : IRealEndPoint
  {
    public const string Type = "Connector" ;

    public static EndPointKey GenerateKey( Connector connector ) => GenerateKey( connector.Owner.Id, connector.Id ) ;
    private static EndPointKey GenerateKey( ElementId equipmentId, int connectorIndex )
    {
      return new EndPointKey( Type, BuildParameterString( equipmentId, connectorIndex, null ) ) ;
    }
    public static string BuildParameterString( Connector connector ) => BuildParameterString( connector.Owner.Id, connector.Id, null ) ;

    internal static ConnectorEndPoint? FromKeyParam( Document document, string param ) => ParseParameterString( document, param ) ;

    private enum SerializeField
    {
      ElementId,
      ConnectorIndex,
      PreferredRadius,
    }

    private static string BuildParameterString( ElementId equipmentId, int connectorIndex, double? preferredRadius )
    {
      var stringifier = new SerializerObject<SerializeField>() ;

      stringifier.Add( SerializeField.ElementId, equipmentId ) ;
      stringifier.Add( SerializeField.ConnectorIndex, connectorIndex ) ;
      stringifier.Add( SerializeField.PreferredRadius, preferredRadius ) ;

      return stringifier.ToString() ;
    }

    public static ConnectorEndPoint? ParseParameterString( Document document, string str )
    {
      var deserializer = new DeserializerObject<SerializeField>( str ) ;

      if ( deserializer.GetElementId( SerializeField.ElementId ) is not { } elementId ) return null ;
      if ( deserializer.GetInt( SerializeField.ConnectorIndex ) is not { } connectorIndex ) return null ;
      var preferredRadius = deserializer.GetDouble( SerializeField.PreferredRadius ) ;

      return new ConnectorEndPoint( document, elementId, connectorIndex, preferredRadius ) ;
    }


    public string TypeName => Type ;
    public string DisplayTypeName => "EndPoint.DisplayTypeName.Connector".GetAppStringByKeyOrDefault( TypeName ) ;

    public EndPointKey Key => GenerateKey( EquipmentId, ConnectorIndex ) ;

    public bool IsReplaceable => true ;

    public bool IsOneSided => true ;

    private readonly Document _document ;

    public ElementId EquipmentId { get ; }
    public int ConnectorIndex { get ; }
    public double? PreferredRadius { get ; }

    public Element? GetOwnerElement() => _document.GetElementById<Instance>( EquipmentId ) ;
    public Connector? GetConnector() => GetOwnerElement()?.GetConnectorManager()?.Lookup( ConnectorIndex ) ;

    public string ParameterString => BuildParameterString( EquipmentId, ConnectorIndex, PreferredRadius ) ;

    public XYZ RoutingStartPosition => GetConnector()?.Origin ?? XYZ.Zero ;

    public ElementId GetLevelId( Document document ) => GetOwnerElement()?.GetLevelId() ?? ElementId.InvalidElementId ;

    public ConnectorEndPoint( Connector connector, double? preferredRadius )
    {
      _document = connector.Owner.Document ;
      PreferredRadius = preferredRadius ;
      EquipmentId = connector.Owner.Id ;
      ConnectorIndex = connector.Id ;
    }

    private ConnectorEndPoint( Document document, ElementId equipmentId, int connectorIndex, double? preferredRadius )
    {
      _document = document ;
      PreferredRadius = preferredRadius ;
      EquipmentId = equipmentId ;
      ConnectorIndex = connectorIndex ;
    }

    public XYZ GetRoutingDirection( bool isFrom )
    {
      return GetConnector()?.CoordinateSystem.BasisZ.ForEndPointType( isFrom ) ?? XYZ.BasisX ;
    }

    public bool HasValidElement( bool isFrom ) => ( null != GetConnector() ) ;

    public Connector? GetReferenceConnector() => GetConnector() ;

    public double? GetDiameter() => PreferredRadius * 2 ?? GetConnector()?.GetDiameter() ;

    public double GetMinimumStraightLength( double edgeDiameter, bool isFrom ) => 0 ;

    Route? IEndPoint.ParentRoute() => null ;
    SubRoute? IEndPoint.ParentSubRoute() => null ;

    public bool GenerateInstance( string routeName ) => false ;
    public bool EraseInstance() => false ;

    public override string ToString() => this.Stringify() ;

    public void Accept( IEndPointVisitor visitor ) => visitor.Visit( this ) ;
    public T Accept<T>( IEndPointVisitor<T> visitor ) => visitor.Visit( this ) ;
  }
}