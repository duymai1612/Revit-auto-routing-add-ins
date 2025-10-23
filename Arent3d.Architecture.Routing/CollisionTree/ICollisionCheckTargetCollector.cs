using System.Collections.Generic ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public interface ICollisionCheckTargetCollector
  {
    BuiltInCategory[] GetCategoriesOfRoutes() ;

    IEnumerable<ElementFilter> CreateElementFilters() ;

    bool IsCollisionCheckElement( Element element ) ;
  }
}