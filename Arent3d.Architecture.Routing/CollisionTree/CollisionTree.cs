using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Reflection ;
using Arent3d.Revit ;
using Arent3d.Routing ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using MathLib ;
using Line = Autodesk.Revit.DB.Line ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public class CollisionTree : ICollisionCheck
  {
    private readonly Document _document ;
    private readonly ICollisionCheckTargetCollector _collector ;
    private readonly IReadOnlyCollection<ElementFilter> _filters ;
    private readonly BuiltInCategory[] _categoriesOnRack ;
    private readonly IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> _routeConditions ;

    public CollisionTree( Document document, ICollisionCheckTargetCollector collector, IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> routeConditions )
    {
      _document = document ;
      _collector = collector ;
      _filters = collector.CreateElementFilters().EnumerateAll() ;
      _categoriesOnRack = collector.GetCategoriesOfRoutes() ;
      _routeConditions = routeConditions ;
    }

    public IEnumerable<Box3d> GetCollidedBoxes( Box3d box ) => Enumerable.Empty<Box3d>() ;

    private static Solid CreateBoundingBoxSolid( XYZ min, XYZ max )
    {
      return GeometryCreationUtilities.CreateExtrusionGeometry( new[] { CreateBaseCurveLoop( min, max ) }, XYZ.BasisZ, max.Z - min.Z ) ;

      static CurveLoop CreateBaseCurveLoop( XYZ min, XYZ max )
      {
        var p1 = min ;
        var p2 = new XYZ( max.X, min.Y, min.Z ) ;
        var p3 = new XYZ( max.X, max.Y, min.Z ) ;
        var p4 = new XYZ( min.X, max.Y, min.Z ) ;
        return CurveLoop.Create( new Curve[] { Line.CreateBound( p1, p2 ), Line.CreateBound( p2, p3 ), Line.CreateBound( p3, p4 ), Line.CreateBound( p4, p1 ) } ) ;
      }
    }

    private static Box3d GetBoundingBox( Element element )
    {
      return element.get_BoundingBox( null ).To3dRaw() ;
    }

    public IEnumerable<(Box3d, IRouteCondition?, bool)> GetCollidedBoxesAndConditions( in Box3d box, CollisionCheckStructureOption option = CollisionCheckStructureOption.CheckAll )
    {
      var min = box.Min.ToXYZRaw() ;
      var max = box.Max.ToXYZRaw() ;

      if ( IsTooSmall( max - min, _document.Application.ShortCurveTolerance ) ) return Enumerable.Empty<(Box3d, IRouteCondition?, bool)>() ;

      var boxFilter = new BoundingBoxIntersectsFilter( new Outline( min, max ) ) ;
      var geometryFilter = new ElementIntersectsSolidFilter( CreateBoundingBoxSolid( min, max ) ) ;
      var elements = _document.GetAllElements<Element>().Where( boxFilter ) ;
      var targets = _filters.Aggregate( elements, ( current, filter ) => current.Where( filter ) ).Where( geometryFilter ).Where( _collector.IsCollisionCheckElement ) ;
      return GetCollidedBoxesAndConditionsImpl( targets ) ;

      static bool IsTooSmall( XYZ size, double tolerance )
      {
        return ( size.X <= tolerance || size.Y <= tolerance || size.Z <= tolerance ) ;
      }
    }

    private IEnumerable<(Box3d, IRouteCondition?, bool)> GetCollidedBoxesAndConditionsImpl( IEnumerable<Element> elements )
    {
      return elements.Select( element => ( GetBoundingBox( element ), GetRouteCondition( element ), false ) ) ;
    }

    private IRouteCondition? GetRouteCondition( Element element )
    {
      if ( Array.IndexOf( _categoriesOnRack, element.GetBuiltInCategory() ) < 0 ) return null ; // Not a routing element.
      if ( element.GetSubRouteInfo() is not { } subRouteInfo ) return null ;  // Not a routing element
      if ( false == _routeConditions.TryGetValue( subRouteInfo, out var routeCondition ) ) return null ;

      return routeCondition ;
    }

    public IEnumerable<(Box3d, IRouteCondition, bool)> GetCollidedBoxesInDetailToRack( Box3d box ) => Enumerable.Empty<(Box3d, IRouteCondition, bool)>() ;

    public IEnumerable<(ElementId ElementId, MeshTriangle Triangle)> GetTriangles()
    {
      var elements = _document.GetAllElements<Element>() ;
      var filteredElements = _filters.Aggregate( elements, ( current, filter ) => current.Where( filter ) ).Where( _collector.IsCollisionCheckElement ) ;
      foreach ( var element in filteredElements.Where( elm => false == elm.IsAutoRoutingGeneratedElement() ) ) {
        var elementId = element.Id ;
        foreach ( var face in element.GetFaces() ) {
          var mesh = face.Triangulate() ;
          for ( int i = 0, n = mesh.NumTriangles ; i < n ; ++i ) {
            yield return ( elementId, mesh.get_Triangle( i ) ) ;
          }
        }
      }
    }
  }

  internal static class SolidExtensions
  {
    public static IEnumerable<Face> GetFaces( this Element element ) => element.GetFineSolids().SelectMany( solid => solid.Faces.OfType<Face>() ) ;

    private static IEnumerable<Solid> GetFineSolids( this Element element ) => element.GetSolids( new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = false, IncludeNonVisibleObjects = false } ) ;

    private static IEnumerable<Solid> GetSolids( this Element element, Options options )
    {
      if ( element.get_Geometry( options ) is not { } geom ) return Enumerable.Empty<Solid>() ;
      return geom.GetSolids() ;
    }

    private static IEnumerable<Solid> GetSolids( this GeometryElement geometry )
    {
      var solids = geometry.OfType<Solid>().Where( solid => false == solid.Faces.IsEmpty ).ToList() ;
      if ( 0 < solids.Count ) return solids ;

      var instanceGeometryElements = geometry.OfType<GeometryInstance>().Select( geom => geom.GetInstanceGeometry() ) ;
      return instanceGeometryElements.SelectMany( GetSolids ) ;
    }
  }
}