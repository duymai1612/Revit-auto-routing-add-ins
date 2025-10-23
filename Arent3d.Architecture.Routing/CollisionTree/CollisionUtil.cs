using System.Collections.Generic ;
using System.Linq ;
using Arent3d.CollisionLib ;
using Arent3d.GeometryLib ;
using Arent3d.Routing ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public static class CollisionUtil 
  {
    private static IEnumerable<IGeometryBody> Intersects( this ITree tree, IGeometryBody target )
    {
      var elements = tree.BoxIntersects( target ) ;
      return 
        GetBodies( elements )
        .Where( intersect => ! AllowIntersection( intersect, target ) ) //トポロジー判定
        .Where( intersect => Arent3d.CollisionLib.GeometryUtil.Intersect( intersect, target ).Any() ) ; //詳細判定
    }
    
    private static IEnumerable<(IGeometryBody iGeometryBody, IGeometry iGeometry)> IntersectsInDetailToRack ( this ITree tree, IGeometryBody target )
    {
      var elements = tree.BoxIntersects( target ) ;
      var tuples = GetBodies( elements )
          .Where( intersect => ! AllowIntersection( intersect, target ) ) //トポロジー判定
          .Select( intersect => new { iGeometryBody = intersect, iGeometries = Arent3d.CollisionLib.GeometryUtil.Intersect( intersect, target ) } ) ;

      foreach ( var tuple in tuples ) {
        foreach ( var iGeometry in tuple.iGeometries ) {
          yield return ( tuple.iGeometryBody, iGeometry ) ;
        }
      }
    }
    
    public static IEnumerable<(IGeometryBody target,IGeometryBody intersect)> IntersectsContinuousBlock( 
      this ITree tree, IReadOnlyCollection<IGeometryBody> continuousTargets )
    {
      var result = 
        from target in continuousTargets 
        let elements = tree.BoxIntersects( target ) 
        from pair in GetBodies( elements )
        .Where( intersect => ! continuousTargets.Contains( intersect ) ) //ターゲット群に含まれるか？
        .Where( intersect => ! AllowIntersection( intersect, target ) ) //トポロジー判定
        .Where( intersect => Arent3d.CollisionLib.GeometryUtil.Intersect( intersect, target ).Any() ) //詳細判定
        .Select( intersect => ( target, intersect ) ) 
        select pair ;
      return result ;
    }

    public static IEnumerable<(IGeometryBody body, IRouteCondition? cond, bool isStructure)>
      GetIntersectAndRoutingCondition( this ITree tree, IGeometryBody target )
    {
      return tree.Intersects( target ).Select( intersect => ( intersect, intersect.ToCondition(), intersect.IsStructure() ) ) ;
    }
    
    public static IEnumerable<(IGeometry body, IRouteCondition? cond )> GetIntersectsInDetailToRack( this ITree tree, IGeometryBody target )
    {
      foreach ( var pair in tree.IntersectsInDetailToRack( target ) ) {
        yield return ( pair.iGeometry, pair.iGeometryBody.ToCondition() ) ;
      }
    }

    private static bool AllowIntersection( IGeometryBody body1, IGeometryBody body2 )
    {
      if ( body1 is FamilyInstanceGeometryBody instanceBody1 && body2 is FamilyInstanceGeometryBody instanceBody2 ) {
        return AreAdjacent( instanceBody1, instanceBody2 ) ;
      }

      return false ;
    }

    private static bool AreAdjacent( FamilyInstanceGeometryBody body1, FamilyInstanceGeometryBody body2 )
    {
      if ( body1.FamilyInstance.GetConnectorManager() is { } manager1 && body2.FamilyInstance.GetConnectorManager() is { } manager2 ) {
        return AreAdjacent( manager1.Connectors, manager2.Connectors ) ;
      }

      return false ;
    }

    private static bool AreAdjacent( ConnectorSet connectors1, ConnectorSet connectors2 )
    {
      var another = connectors2.OfType<Connector>().OfEnd().Select( EndPoints.ConnectorEndPoint.GenerateKey ).ToHashSet() ;
      return connectors1.OfType<Connector>().SelectMany( c => c.GetConnectedConnectors().OfEnd() ).Any( c => another.Contains( EndPoints.ConnectorEndPoint.GenerateKey( c ) ) ) ;
    }

    private static IRouteCondition? ToCondition( this IGeometryBody body )
    {
      if ( ! ( body is FamilyInstanceGeometryBody geometryBody ) ) return null ;

      return geometryBody.GetRouteCondition() ;
    }

    private static bool IsStructure( this IGeometryBody body )
    {
      if ( ! ( body is FamilyInstanceGeometryBody geometryBody ) ) return false ;
      
      return geometryBody.IsStructure ;
    }

    private static IEnumerable<IGeometryBody> GetBodies( IEnumerable<TreeElement> elements)
    {
      var cash = new HashSet<IGeometryBody>() ;
      foreach(var element in elements) {
        switch ( element.Value ) {
          case IGeometryBody b :
            if ( ! cash.Add( b ) )
              continue ;
            yield return b ;
            break ;
          case ComplexGeometryBodyPart part :
            var comp = part.ParentBody ;
            if ( ! cash.Add( comp ) ) 
              continue ;
            yield return comp ;
            break ;
          default :
            continue ;
        }
      }
    }
    
    public static GeometryElement? GetGeometryElement( this Element element )
    {
      return element.get_Geometry( new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = false, IncludeNonVisibleObjects = false } ) ;
    }
  }
}

