using System.Collections.Generic ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public class PipingCollisionCheckTargetCollector : CollisionCheckTargetCollectorBase
  {
    public PipingCollisionCheckTargetCollector( Document document, IReadOnlyCollection<Route> routes ) : base( document )
    {
    }

    public override BuiltInCategory[] GetCategoriesOfRoutes() => (BuiltInCategory[])BuiltInCategorySets.Pipes.Clone() ;

    public override bool IsCollisionCheckElement( Element elm )
    {
      if ( elm is not FamilyInstance fi ) return true ;

      // Racks are not collision targets.
      return fi.IsFamilyInstanceExcept( RoutingFamilyType.PassPoint, RoutingFamilyType.RackGuide ) ;
    }
  }
}