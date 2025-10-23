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
  public class PassPointBranchEndPoint : IRouteBranchEndPoint, IEndPointOfPassPoint
  {
    public const string Type = "Pass Point Branch" ;

    private enum SerializeField
    {
      PassPointId,
      Diameter,
      EndPointKeyOverSubRoute,
    }

    public static PassPointBranchEndPoint? ParseParameterString( Document document, string str )
    {
      var deserializer = new DeserializerObject<SerializeField>( str ) ;

      if ( deserializer.GetElementId( SerializeField.PassPointId ) is not { } passPointId ) return null ;
      var diameter = deserializer.GetDouble( SerializeField.Diameter ) ;
      if ( deserializer.GetEndPointKey( SerializeField.EndPointKeyOverSubRoute ) is not { } referenceEndPointKey ) return null ;

      return new PassPointBranchEndPoint( document, passPointId, diameter * 0.5, referenceEndPointKey ) ;
    }

    public string ParameterString
    {
      get
      {
        var stringifier = new SerializerObject<SerializeField>() ;

        stringifier.Add( SerializeField.PassPointId, PassPointId ) ;
        stringifier.Add( SerializeField.Diameter, GetDiameter() ) ;
        stringifier.AddNonNull( SerializeField.EndPointKeyOverSubRoute, EndPointKeyOverSubRoute ) ;

        return stringifier.ToString() ;
      }
    }


    public string TypeName => Type ;
    public string DisplayTypeName => "EndPoint.DisplayTypeName.PassPointBranch".GetAppStringByKeyOrDefault( TypeName ) ;
    public EndPointKey Key => new EndPointKey( TypeName, PassPointId.IntegerValue.ToString() ) ;

    public EndPointKey EndPointKeyOverSubRoute { get ; }

    public bool IsReplaceable => false ;

    public bool IsOneSided => false ;

    internal Document Document { get ; }

    public ElementId PassPointId { get ; }

    public Instance? GetPassPoint() => Document.GetElementById<Instance>( PassPointId ) ;

    public string RouteName => GetPassPoint()?.GetRouteName() ?? string.Empty ;

    public SubRoute? GetSubRoute( bool fromSideOfPassPoint )
    {
      if ( ParentRoute() is not { } route ) return null ;

      var passPointKey = PassPointEndPoint.KeyFromPassPointId( PassPointId ) ;
      return route.SubRoutes.FirstOrDefault( subRoute => HasPassPoint( subRoute, passPointKey, fromSideOfPassPoint ) ) ;

      static bool HasPassPoint( SubRoute subRoute, EndPointKey passPointKey, bool fromSideOfPassPoint )
      {
        return subRoute.Segments.Any( s => ( fromSideOfPassPoint ? s.ToEndPoint : s.FromEndPoint ).Key == passPointKey ) ;
      }
    }

    public XYZ RoutingStartPosition => GetPreferredStartPosition() ?? XYZ.Zero ;
    public XYZ Direction => GetDirectionFromAngle() ?? XYZ.BasisX ;

    private XYZ? GetDirectionFromAngle()
    {
      if ( GetPassPoint() is not { } passPoint ) return null ;

      var transform = passPoint.GetTotalTransform() ;
      return transform.BasisX ;
    }

    private double? PreferredRadius { get ; set ; }

    public ElementId GetLevelId( Document document ) => GetPassPoint()?.GetLevelId() ?? ElementId.InvalidElementId ;

    public void UpdatePreferredParameters()
    {
      if ( GetPassPoint() is not { } passPoint ) return ;

      SetPreferredParameters( passPoint ) ;
    }

    private void SetPreferredParameters( Instance passPoint )
    {
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

    private static double? GetForcedFixedHeight( Document document, FixedHeight? fixedHeight, ElementId levelId )
    {
      if ( null == fixedHeight ) return null ;

      return document.GetHeightSettingStorable().GetAbsoluteHeight( levelId, fixedHeight.Value.Type, fixedHeight.Value.Height ) ;
    }

    private static IEnumerable<RouteSegment> GetRelatedSegments( Route route, EndPointKey pointKey )
    {
      return route.RouteSegments.Where( s => s.FromEndPoint.Key == pointKey || s.ToEndPoint.Key == pointKey ) ;
    }

    public PassPointBranchEndPoint( Document document, ElementId passPointId, double? preferredRadius, EndPointKey endPointKeyOverSubRoute )
    {
      Document = document ;
      PassPointId = passPointId ;

      PreferredRadius = preferredRadius ;

      EndPointKeyOverSubRoute = endPointKeyOverSubRoute ;

      UpdatePreferredParameters() ;
    }

    public XYZ GetRoutingDirection( bool isFrom ) => Direction ;

    public bool HasValidElement( bool isFrom ) => ( null != GetPassPoint() ) ;

    public Connector? GetReferenceConnector() => null ;

    public double? GetDiameter() => GetPassPoint()?.LookupParameter( "Arent-RoundDuct-Diameter" )?.AsDouble() ?? PreferredRadius * 2 ;

    public double GetMinimumStraightLength( double edgeDiameter, bool isFrom ) => 0 ;

    public Route? ParentRoute() => RouteCache.Get( Document ).TryGetValue( RouteName, out var route ) ? route : null ;
    SubRoute? IEndPoint.ParentSubRoute() => null ;

    public bool GenerateInstance( string routeName )
    {
      if ( null == GetPassPoint() ) throw new InvalidOperationException() ;

      return true ;
    }

    public bool EraseInstance() => false ;

    public override string ToString() => this.Stringify() ;

    public void Accept( IEndPointVisitor visitor ) => visitor.Visit( this ) ;
    public T Accept<T>( IEndPointVisitor<T> visitor ) => visitor.Visit( this ) ;
  }
}