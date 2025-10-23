using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;
using Arent3d.Utility.Serialization ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public class RouteSegment
  {
    public MEPSystemClassificationInfo SystemClassificationInfo { get ; set ; }
    public MEPSystemType? SystemType { get ; set ; }
    public MEPCurveType? CurveType { get ; set ; }

    public double? PreferredNominalDiameter { get ; private set ; }

    public FixedHeight? FromFixedHeight { get ; set ; }
    public FixedHeight? ToFixedHeight { get ; set ; }
    public AvoidType AvoidType { get ; set ; }

    public IEndPoint FromEndPoint { get ; private set ; }
    public IEndPoint ToEndPoint { get ; private set ; }
    public bool IsRoutingOnPipeSpace { get ; internal set ; } = false ;

    public ElementId ShaftElementId { get ; internal set ; }

    public IReadOnlyCollection<SubRouteInfo> SubRouteGroup { get ; private set ; } = Array.Empty<SubRouteInfo>() ;

    internal void SetSubRouteGroup( IReadOnlyCollection<SubRouteInfo> subRouteGroup )
    {
      if ( 1 < subRouteGroup.Count ) {
        SubRouteGroup = subRouteGroup ;
      }
      else {
        SubRouteGroup = Array.Empty<SubRouteInfo>() ;
      }
    }

    public double? GetRealNominalDiameter()
    {
      if ( PreferredNominalDiameter.HasValue ) return PreferredNominalDiameter.Value ;

      return GetRealNominalDiameterFromEndPoints() ;
    }

    private double? GetRealNominalDiameterFromEndPoints()
    {
      if ( FromEndPoint.GetDiameter() is { } d1 ) {
        return d1 ;
      }

      if ( ToEndPoint.GetDiameter() is { } d2 ) {
        return d2 ;
      }

      return null ;
    }

    public void ChangePreferredNominalDiameter( double d )
    {
      PreferredNominalDiameter = d ;
    }

    public bool ApplyRealNominalDiameter()
    {
      if ( GetRealNominalDiameterFromEndPoints() is not { } d ) return false ;

      PreferredNominalDiameter = d ;
      return true ;
    }

    public RouteSegment( MEPSystemClassificationInfo classificationInfo, MEPSystemType? systemType, MEPCurveType? curveType, IEndPoint fromEndPoint, IEndPoint toEndPoint, double? preferredNominalDiameter, bool isRoutingOnPipeSpace, FixedHeight? fromFixedHeight, FixedHeight? toFixedHeight, AvoidType avoidType, ElementId shaftElementId )
    {
      SystemClassificationInfo = classificationInfo ;
      SystemType = systemType ;
      CurveType = curveType ;

      PreferredNominalDiameter = ( 0 < preferredNominalDiameter ? preferredNominalDiameter : null ) ;
      IsRoutingOnPipeSpace = isRoutingOnPipeSpace ;
      FromFixedHeight = fromFixedHeight ;
      ToFixedHeight = toFixedHeight ;
      AvoidType = avoidType ;
      FromEndPoint = fromEndPoint ;
      ToEndPoint = toEndPoint ;
      ShaftElementId = shaftElementId ;
    }

    public void ReplaceEndPoint( IEndPoint oldEndPoint, IEndPoint newEndPoint )
    {
      if ( false == oldEndPoint.IsReplaceable ) throw new InvalidOperationException() ;

      if ( FromEndPoint == oldEndPoint ) {
        FromEndPoint = newEndPoint ;
      }

      if ( ToEndPoint == oldEndPoint ) {
        ToEndPoint = newEndPoint ;
      }
    }
  }


  [StorableConverterOf( typeof( RouteSegment ) )]
  internal class RouteSegmentConverter : StorableConverterBase<RouteSegment>
  {
    private enum SerializeField
    {
      PreferredNominalDiameter,
      FromEndPoint,
      ToEndPoint,
      IsRoutingOnPipeSpace,
      CurveType,
      FromFixedBopHeightType,
      FromFixedBopHeight,
      ToFixedBopHeightType,
      ToFixedBopHeight,
      AvoidType,
      SystemClassificationInfo,
      SystemType,
      SubRouteGroup,
      ShaftElementId,
    }

    protected override RouteSegment Deserialize( Element storedElement, IDeserializerObject deserializerObject )
    {
      var deserializer = deserializerObject.Of<SerializeField>() ;

      var preferredDiameter = deserializer.GetDouble( SerializeField.PreferredNominalDiameter ) ;
      var fromId = storedElement.Document.ParseEndPoint( deserializer.GetString( SerializeField.FromEndPoint ) ?? throw new InvalidOperationException() ) ?? throw new InvalidOperationException() ;
      var toId = storedElement.Document.ParseEndPoint( deserializer.GetString( SerializeField.ToEndPoint ) ?? throw new InvalidOperationException() ) ?? throw new InvalidOperationException() ;
      var isRoutingOnPipeSpace = deserializer.GetBool( SerializeField.IsRoutingOnPipeSpace ) ?? throw new InvalidOperationException() ;
      var curveType = deserializer.GetElement<SerializeField, MEPCurveType>( SerializeField.CurveType, storedElement.Document ) ?? throw new InvalidOperationException() ;
      var fromFixedHeightType = deserializer.GetEnum<FixedHeightType>( SerializeField.FromFixedBopHeightType ) ;
      var fromFixedHeightValue = deserializer.GetDouble( SerializeField.FromFixedBopHeight ) ;
      var toFixedHeightType = deserializer.GetEnum<FixedHeightType>( SerializeField.ToFixedBopHeightType ) ;
      var toFixedHeightValue = deserializer.GetDouble( SerializeField.ToFixedBopHeight ) ;
      var avoidType = deserializer.GetEnum<AvoidType>( SerializeField.AvoidType ) ?? throw new InvalidOperationException() ;
      var classificationInfo = MEPSystemClassificationInfo.Deserialize( deserializer.GetString( SerializeField.SystemClassificationInfo ) ?? throw new InvalidOperationException() ) ?? throw new InvalidOperationException() ;
      MEPSystemType? systemType = null ;
      if ( classificationInfo.HasSystemType() ) {
        systemType = deserializer.GetElement<SerializeField, MEPSystemType>( SerializeField.SystemType, storedElement.Document ) ?? throw new InvalidOperationException() ;
      }

      var subRouteGroups = deserializer.GetNonNullArray( SerializeField.SubRouteGroup, SubRouteInfo.CreateForDeserialize ) ;
      var shaftElementId = deserializer.GetElementId( SerializeField.ShaftElementId ) ?? ElementId.InvalidElementId ;

      var fromFixedHeight = FixedHeight.CreateOrNull( fromFixedHeightType, fromFixedHeightValue ) ;
      var toFixedHeight = FixedHeight.CreateOrNull( toFixedHeightType, toFixedHeightValue ) ;

      var routeSegment = new RouteSegment( classificationInfo, systemType, curveType, fromId, toId, preferredDiameter, isRoutingOnPipeSpace, fromFixedHeight, toFixedHeight, avoidType, shaftElementId ) ;
      if ( null != subRouteGroups ) {
        routeSegment.SetSubRouteGroup( subRouteGroups ) ;
      }

      return routeSegment ;
    }

    protected override ISerializerObject Serialize( Element storedElement, RouteSegment customTypeValue )
    {
      var serializerObject = new SerializerObject<SerializeField>() ;

      serializerObject.Add( SerializeField.PreferredNominalDiameter, customTypeValue.PreferredNominalDiameter ) ;
      serializerObject.AddNonNull( SerializeField.FromEndPoint, customTypeValue.FromEndPoint.ToString() ) ;
      serializerObject.AddNonNull( SerializeField.ToEndPoint, customTypeValue.ToEndPoint.ToString() ) ;
      serializerObject.Add( SerializeField.IsRoutingOnPipeSpace, customTypeValue.IsRoutingOnPipeSpace ) ;
      serializerObject.Add( SerializeField.CurveType, customTypeValue.CurveType ) ;
      serializerObject.Add( SerializeField.FromFixedBopHeightType, customTypeValue.FromFixedHeight?.Type ) ;
      serializerObject.Add( SerializeField.FromFixedBopHeight, customTypeValue.FromFixedHeight?.Height ) ;
      serializerObject.Add( SerializeField.ToFixedBopHeightType, customTypeValue.ToFixedHeight?.Type ) ;
      serializerObject.Add( SerializeField.ToFixedBopHeight, customTypeValue.ToFixedHeight?.Height ) ;
      serializerObject.Add( SerializeField.AvoidType, customTypeValue.AvoidType ) ;
      serializerObject.AddNonNull( SerializeField.SystemClassificationInfo, customTypeValue.SystemClassificationInfo.Serialize() ) ;
      serializerObject.Add( SerializeField.SystemType, customTypeValue.SystemType ) ;

      if ( 1 < customTypeValue.SubRouteGroup.Count ) {
        serializerObject.AddNonNull( SerializeField.SubRouteGroup, customTypeValue.SubRouteGroup ) ;
      }

      serializerObject.Add( SerializeField.ShaftElementId, customTypeValue.ShaftElementId ) ;

      return serializerObject ;
    }
  }
}