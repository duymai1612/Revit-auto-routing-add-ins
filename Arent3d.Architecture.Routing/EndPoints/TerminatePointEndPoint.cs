using System.Diagnostics ;
using System.Linq ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.EndPoints
{
  [DebuggerDisplay("{Key}")]
  public class TerminatePointEndPoint : IRealEndPoint
  {
    public const string Type = "Terminate Point" ;

    private enum SerializeField
    {
      TerminatePointId,
      Diameter,
      Position,
      Direction,
      LinkedInstanceId,
    }

    public static TerminatePointEndPoint? ParseParameterString( Document document, string str )
    {
      var deserializer = new DeserializerObject<SerializeField>( str ) ;

      if ( deserializer.GetElementId( SerializeField.TerminatePointId ) is not { } terminatePointId ) return null ;
      var diameter = deserializer.GetDouble( SerializeField.Diameter ) ;
      if ( deserializer.GetXYZ( SerializeField.Position ) is not { } position ) return null ;
      if ( deserializer.GetXYZ( SerializeField.Direction ) is not { } direction ) return null ;
      if ( deserializer.GetElementId( SerializeField.LinkedInstanceId ) is not { } linkedInstanceId ) return null ;

      return new TerminatePointEndPoint( document, terminatePointId, position, direction, diameter * 0.5, linkedInstanceId ) ;
    }

    public string ParameterString
    {
      get
      {
        var stringifier = new SerializerObject<SerializeField>() ;

        stringifier.Add( SerializeField.TerminatePointId, TerminatePointId ) ;
        stringifier.Add( SerializeField.Diameter, GetDiameter() ) ;
        stringifier.AddNonNull( SerializeField.Position, RoutingStartPosition ) ;
        stringifier.AddNonNull( SerializeField.Direction, Direction ) ;
        stringifier.Add( SerializeField.LinkedInstanceId, LinkedInstanceId ) ;

        return stringifier.ToString() ;
      }
    }


    public string TypeName => Type ;
    public string DisplayTypeName => "EndPoint.DisplayTypeName.Terminal".GetAppStringByKeyOrDefault( TypeName ) ;
    public EndPointKey Key => new EndPointKey( TypeName, TerminatePointId.IntegerValue.ToString() ) ;

    internal static TerminatePointEndPoint? FromKeyParam( Document document, string param )
    {
      if ( false == int.TryParse( param, out var terminatePointId ) ) return null ;
      if ( document.GetElementById<FamilyInstance>( terminatePointId ) is not { } instance ) return null ;
      if ( instance.Symbol.Id != document.GetFamilySymbols( RoutingFamilyType.TerminatePoint ).FirstOrDefault()?.Id ) return null ;

      return new TerminatePointEndPoint( instance, null ) ;
    }

    public bool IsReplaceable => true ;

    public bool IsOneSided => true ;

    private readonly Document _document ;

    public ElementId TerminatePointId { get ; private set ; }
    public ElementId LinkedInstanceId { get ; private set ; }

    public Instance? GetTerminatePoint() => _document.GetElementById<Instance>( TerminatePointId ) ;

    private XYZ PreferredPosition { get ; set ; } = XYZ.Zero ;

    public XYZ RoutingStartPosition => GetTerminatePoint()?.GetTotalTransform().Origin ?? PreferredPosition ;
    private XYZ PreferredDirection { get ; set ; } = XYZ.Zero ;
    public XYZ Direction => GetTerminatePoint()?.GetTotalTransform().BasisX ?? PreferredDirection ;
    private double? PreferredRadius { get ; set ; } = 0 ;

    public ElementId GetLevelId( Document document ) => GetTerminatePoint()?.GetLevelId() ?? GetElementLevelId( document, LinkedInstanceId ) ?? document.GuessLevelId( PreferredPosition ) ;

    private static ElementId? GetElementLevelId( Document document, ElementId linkedInstanceId )
    {
      return document.GetElementById<Element>( linkedInstanceId )?.GetLevelId() ;
    }

    public void UpdatePreferredParameters()
    {
      if ( GetTerminatePoint() is not { } terminatePoint ) return ;

      SetPreferredParameters( terminatePoint ) ;
    }

    private void SetPreferredParameters( Instance terminatePoint )
    {
      var transform = terminatePoint.GetTotalTransform() ;
      PreferredPosition = transform.Origin ;
      PreferredDirection = transform.BasisX ;
      PreferredRadius = terminatePoint.LookupParameter( "Arent-RoundDuct-Diameter" )?.AsDouble() * 0.5 ;
    }

    public TerminatePointEndPoint( Instance instance, Instance? linkedInstance )
    {
      _document = instance.Document ;
      TerminatePointId = instance.Id ;
      LinkedInstanceId = linkedInstance.GetValidId() ;

      SetPreferredParameters( instance ) ;
    }

    public TerminatePointEndPoint( Document document, ElementId terminatePointId, XYZ preferredPosition, XYZ preferredDirection, double? preferredRadius, ElementId linkedInstanceId )
    {
      _document = document ;
      TerminatePointId = terminatePointId ;
      LinkedInstanceId = linkedInstanceId ;

      PreferredPosition = preferredPosition ;
      PreferredDirection = ( preferredDirection.IsZeroLength() ? XYZ.BasisX : preferredDirection.Normalize() ) ;
      PreferredRadius = preferredRadius ;
      UpdatePreferredParameters() ;
    }

    public XYZ GetRoutingDirection( bool isFrom ) => Direction ;

    public bool HasValidElement( bool isFrom ) => ( null != GetTerminatePoint() ) ;

    public Connector? GetReferenceConnector() => null ;

    public double? GetDiameter() => GetTerminatePoint()?.LookupParameter( "Arent-RoundDuct-Diameter" )?.AsDouble() ?? PreferredRadius * 2 ;

    public double GetMinimumStraightLength( double edgeDiameter, bool isFrom ) => 0 ;

    Route? IEndPoint.ParentRoute() => null ;
    SubRoute? IEndPoint.ParentSubRoute() => null ;

    public bool GenerateInstance( string routeName )
    {
      if ( null != GetTerminatePoint() ) return false ;

      TerminatePointId = _document.AddTerminatePoint( routeName, PreferredPosition, PreferredDirection, PreferredRadius, GetLevelId( _document ) ).Id ;

      Element elemTerP = _document.GetElement( TerminatePointId ) ;
      Element elemOrg = _document.GetElement( LinkedInstanceId ) ;

      foreach ( Parameter parameter in elemTerP.Parameters ) {
        if ( parameter.Definition.Name == "LinkedInstanceId" ) {
          parameter.Set( LinkedInstanceId.ToString() ) ;
        }

        if ( parameter.Definition.Name == "LinkedInstanceXYZ" ) {
          LocationPoint? Lp = elemOrg.Location as LocationPoint ;
          XYZ? ElementPoint = Lp?.Point ;
          XYZ addPoint = PreferredPosition - ElementPoint ;

          parameter.Set( addPoint.ToString().Substring( 1, addPoint.ToString().Length - 1 ) ) ;
        }
      }

      elemTerP.SetProperty( PassPointParameter.RelatedConnectorId, LinkedInstanceId.IntegerValue.ToString() ) ;
      elemTerP.SetProperty( PassPointParameter.RelatedFromConnectorId, LinkedInstanceId.IntegerValue.ToString() ) ;

      return true ;
    }

    public bool EraseInstance()
    {
      UpdatePreferredParameters() ;
      return ( 0 < _document.Delete( TerminatePointId ).Count ) ;
    }

    public override string ToString() => this.Stringify() ;

    public void Accept( IEndPointVisitor visitor ) => visitor.Visit( this ) ;
    public T Accept<T>( IEndPointVisitor<T> visitor ) => visitor.Visit( this ) ;
  }
}