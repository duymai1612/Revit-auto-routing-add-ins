using System ;
using System.Collections.Generic ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public abstract class CollisionCheckTargetCollectorBase : ICollisionCheckTargetCollector
  {
    private readonly Document _document ;
    
    protected CollisionCheckTargetCollectorBase( Document document )
    {
      _document = document ;
    }

    public abstract BuiltInCategory[] GetCategoriesOfRoutes() ;

    public IEnumerable<ElementFilter> CreateElementFilters()
    {
      yield return new ElementMulticategoryFilter( BuiltInCategorySets.Obstacles ) ;
    }

    public abstract bool IsCollisionCheckElement( Element element ) ;
  }
}