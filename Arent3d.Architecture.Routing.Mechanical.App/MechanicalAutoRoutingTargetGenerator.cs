using System.Collections.Generic ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Mechanical.App
{
  public class MechanicalAutoRoutingTargetGenerator : AutoRoutingTargetGenerator
  {
    public MechanicalAutoRoutingTargetGenerator( Document document ) : base( document )
    {
    }

    protected override IReadOnlyCollection<AutoRoutingTarget> GenerateAutoRoutingTarget( IReadOnlyCollection<SubRoute> subRoutes, IReadOnlyDictionary<Route, int> priorities, IReadOnlyDictionary<SubRouteInfo, MEPSystemRouteCondition> routeConditionDictionary )
    {
      return new[] { new AutoRoutingTarget( Document, subRoutes, priorities, routeConditionDictionary ) } ;
    }
  }
}