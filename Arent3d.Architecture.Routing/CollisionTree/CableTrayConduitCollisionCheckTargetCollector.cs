using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public class CableTrayConduitCollisionCheckTargetCollector : CollisionCheckTargetCollectorBase
  {
    public CableTrayConduitCollisionCheckTargetCollector( Document document, IReadOnlyCollection<Route> routes ) : base( document )
    {
    }

    public override BuiltInCategory[] GetCategoriesOfRoutes() => (BuiltInCategory[])BuiltInCategorySets.ElectricalRoutingElements.Clone() ;

    public override bool IsCollisionCheckElement( Element elm )
    {
      if ( IsConduit( elm ) ) return false ;
      
      if ( elm is not FamilyInstance fi ) return true ;

      // Racks are collision targets.
      return fi.IsFamilyInstanceExcept( RoutingFamilyType.PassPoint, RoutingFamilyType.Shaft, RoutingFamilyType.RackSpace ) ;
    }

    private static bool IsConduit( Element elm )
    {
      return BuiltInCategorySets.ElectricalRoutingElements.Contains( elm.GetBuiltInCategory() ) ;
    }
  }
}