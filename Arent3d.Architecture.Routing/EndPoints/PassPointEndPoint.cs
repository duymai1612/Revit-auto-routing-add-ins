using System ;
using System.Collections.Generic ;
using System.Diagnostics ;
using System.Linq ;
using Arent3d.Architecture.Routing.Extensions ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.EndPoints
{
  [DebuggerDisplay("{Key}")]
  public class PassPointEndPoint : IEndPointOfPassPoint
  {
    public const string Type = "Pass Point" ;

    private enum SerializeField
    {
      PassPointId,
      Diameter,
      Position,
      Direction,
      Level,
    }

    public static PassPointEndPoint? ParseParameterString( Document document, string str )
    {
      var deserializer = new DeserializerObject<SerializeField>( str ) ;

      if ( deserializer.GetElementId( SerializeField.PassPointId ) is not { } passPointId ) return null ;
      var diameter = deserializer.GetDouble( SerializeField.Diameter ) ;
      if ( deserializer.GetXYZ( SerializeField.Position ) is not { } position ) return null ;
      if ( deserializer.GetXYZ( SerializeField.Direction ) is not { } direction ) return null ;

      Level? level = null ;
      if ( deserializer.GetString( SerializeField.Level ) is { } levelName ) {
        level = document.GetAllElements<Level>().FirstOrDefault( elm => elm.Name == levelName ) ;
      }

      return new PassPointEndPoint( document, passPointId, position, direction, diameter * 0.5, level ) ;
    }

    public string ParameterString
    {
      get
      {
        var stringifier = new SerializerObject<SerializeField>() ;

        stringifier.Add( SerializeField.PassPointId, PassPointId ) ;
        stringifier.Add( SerializeField.Diameter, GetDiameter() ) ;
        stringifier.AddNonNull( SerializeField.Position, RoutingStartPosition ) ;
        stringifier.AddNonNull( SerializeField.Direction, Direction ) ;
        stringifier.AddNullable( SerializeField.Level, Document.GetElementById<Level>( PreferredLevelId )?.Name ) ;

        return stringifier.ToString() ;
      }
    }


    public string TypeName => Type ;
    public string DisplayTypeName => "EndPoint.DisplayTypeName.PassPoint".GetAppStringByKeyOrDefault( TypeName ) ;

    public EndPointKey Key => KeyFromPassPointId( PassPointId ) ;

    public static EndPointKey KeyFromPassPointId( ElementId passPointId ) => new EndPointKey( Type, passPointId.IntegerValue.ToString() ) ;

    internal static PassPointEndPoint? FromKeyParam( Document document, string param )
    {
      if ( false == int.TryParse( param, out var passPointId ) ) return null ;
      if ( document.GetElementById<FamilyInstance>( passPointId ) is not { } instance ) return null ;
      if ( instance.Symbol.Id != document.GetFamilySymbols( RoutingFamilyType.PassPoint ).FirstOrDefault()?.Id ) return null ;

      return new PassPointEndPoint( instance ) ;
    }

    public bool IsReplaceable => false ;

    public bool IsOneSided => false ;

    internal Document Document { get ; }

    public ElementId PassPointId { get ; private set ; }

    public Instance? GetPassPoint() => Document.GetElementById<Instance>( PassPointId ) ;

    private XYZ PreferredPosition { get ; set ; } = XYZ.Zero ;

    public XYZ RoutingStartPosition => GetPreferredStartPosition() ?? PreferredPosition ;
    private XYZ PreferredDirection { get ; set ; } = XYZ.Zero ;
    public XYZ Direction => GetPassPoint()?.GetTotalTransform().BasisX ?? PreferredDirection ;
    private double? PreferredRadius { get ; set ; } = 0 ;
    public ElementId PreferredLevelId { get ; }

    public ElementId GetLevelId( Document document ) => GetPassPoint()?.GetLevelId() ?? ( ElementId.InvalidElementId != PreferredLevelId ? PreferredLevelId : document.GuessLevelId( PreferredPosition ) ) ;

    public void UpdatePreferredParameters()
    {
      if ( GetPassPoint() is not { } passPoint ) return ;

      SetPreferredParameters( passPoint ) ;
    }

    private void SetPreferredParameters( Instance passPoint )
    {
      var transform = passPoint.GetTotalTransform() ;
      PreferredPosition = transform.Origin ;
      PreferredDirection = transform.BasisX ;
      PreferredRadius = passPoint.LookupParameter( "Arent-RoundDuct-Diameter" )?.AsDouble() * 0.5 ;
    }

    /// <summary>
    /// return start position after comparing to FixedBopHeight
    /// </summary>
    /// <returns></returns>
    private XYZ? GetPreferredStartPosition()
    {
      if ( GetPassPoint() is not { } passPoint ) return null ;

      var startPosition = passPoint.GetTotalTransform().Origin ;

      if ( passPoint.GetRouteName() is not { } routeName || false == RouteCache.Get( passPoint.Document ).TryGetValue( routeName, out var route ) ) return startPosition ;

      var segments = GetRelatedSegments( route, Key ) ;
      var document = passPoint.Document ;
      var targetLevelId = passPoint.GetLevelId() ;
      var segmentsAndFixedHeights = segments.Select( s => ( Segment: s, FixedHeight: GetForcedFixedHeight( document, targetLevelId, s ) ) ) ;

      foreach ( var (targetSegment, fixedHeight) in segmentsAndFixedHeights.Where( tuple => tuple.FixedHeight.HasValue ) ) {
        if ( targetSegment.PreferredNominalDiameter is not { } diameter ) break ;

        var fixedCenterHeight = fixedHeight!.Value ;
        var passPointZ = passPoint.GetTotalTransform().Origin.Z ;
        var difference = Math.Abs( fixedCenterHeight - passPointZ ) ;
        if ( diameter <= difference ) break ;

        return new XYZ( startPosition.X, startPosition.Y, fixedCenterHeight ) ;
      }

      return startPosition ;
    }


    private static double? GetForcedFixedHeight( Document document, ElementId levelId, RouteSegment segment )
    {
      if ( null == segment.FromFixedHeight && null == segment.ToFixedHeight || ElementId.InvalidElementId == levelId ) return null ;

      if ( segment.FromEndPoint.GetLevelId( document ) == levelId ) return GetForcedFixedHeight( document, segment.FromFixedHeight, levelId ) ;
      if ( segment.ToEndPoint.GetLevelId( document ) == levelId ) return GetForcedFixedHeight( document, segment.ToFixedHeight, levelId ) ;

      return null ;
    }

    public static double? GetForcedFixedHeight( Document document, FixedHeight? fixedHeight, ElementId levelId )
    {
      if ( null == fixedHeight ) return null ;

      return document.GetHeightSettingStorable().GetAbsoluteHeight( levelId, fixedHeight.Value.Type, fixedHeight.Value.Height ) ;
    }

    private static IEnumerable<RouteSegment> GetRelatedSegments( Route route, EndPointKey pointKey )
    {
      return route.RouteSegments.Where( s => s.FromEndPoint.Key == pointKey || s.ToEndPoint.Key == pointKey ) ;
    }

    public PassPointEndPoint( Instance instance )
    {
      Document = instance.Document ;
      PassPointId = instance.Id ;
      PreferredLevelId = instance.GetLevelId() ;

      SetPreferredParameters( instance ) ;
    }

    public PassPointEndPoint( Document document, ElementId passPointId, XYZ preferredPosition, XYZ preferredDirection, double? preferredRadius, Level? level )
    {
      Document = document ;
      PassPointId = passPointId ;

      PreferredPosition = preferredPosition ;
      PreferredDirection = ( preferredDirection.IsZeroLength() ? XYZ.BasisX : preferredDirection.Normalize() ) ;
      PreferredRadius = preferredRadius ;
      PreferredLevelId = level?.LevelId ?? ElementId.InvalidElementId ;
      UpdatePreferredParameters() ;
    }

    public XYZ GetRoutingDirection( bool isFrom ) => Direction ;

    public bool HasValidElement( bool isFrom ) => ( null != GetPassPoint() ) ;

    public Connector? GetReferenceConnector() => null ;

    public double? GetDiameter() => GetPassPoint()?.LookupParameter( "Arent-RoundDuct-Diameter" )?.AsDouble() ?? PreferredRadius * 2 ;

    public double GetMinimumStraightLength( double edgeDiameter, bool isFrom ) => 0 ;

    Route? IEndPoint.ParentRoute() => null ;
    SubRoute? IEndPoint.ParentSubRoute() => null ;

    public bool GenerateInstance( string routeName )
    {
      if ( null != GetPassPoint() ) return true ;

      PassPointId = Document.AddPassPoint( routeName, PreferredPosition, PreferredDirection, PreferredRadius, PreferredLevelId ).Id ;
      return false ;
    }

    public bool EraseInstance()
    {
      UpdatePreferredParameters() ;
      return ( 0 < Document.Delete( PassPointId ).Count ) ;
    }

    public override string ToString() => this.Stringify() ;

    public void Accept( IEndPointVisitor visitor ) => visitor.Visit( this ) ;
    public T Accept<T>( IEndPointVisitor<T> visitor ) => visitor.Visit( this ) ;
  }
}