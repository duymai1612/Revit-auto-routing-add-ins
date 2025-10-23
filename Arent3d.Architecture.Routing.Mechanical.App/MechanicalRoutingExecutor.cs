using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.AppBase ;
using Arent3d.Architecture.Routing.CollisionTree ;
using Arent3d.Architecture.Routing.FittingSizeCalculators ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Mechanical.App
{
  public class MechanicalRoutingExecutor : RoutingExecutor
  {
    public MechanicalRoutingExecutor( Document document, View view, IFittingSizeCalculator fittingSizeCalculator ) : base( document, view, fittingSizeCalculator )
    {
    }

    protected override IEnumerable<FamilyInstance> GetRackFamilyInstances()
    {
      return Document.GetAllFamilyInstances( RoutingFamilyType.RackGuide ) ;
    }

    protected override RouteGenerator CreateRouteGenerator( IReadOnlyCollection<Route> routes, Document document, ICollisionCheckTargetCollector collector )
    {
      return new RouteGenerator( document, routes, new MechanicalAutoRoutingTargetGenerator( document ), FittingSizeCalculator, collector ) ;
    }

    protected override ICollisionCheckTargetCollector CreateCollisionCheckTargetCollector( Domain domain, IReadOnlyCollection<Route> routesInType )
    {
      return domain switch
      {
        Domain.DomainHvac => new HVacCollisionCheckTargetCollector( Document, routesInType ),
        Domain.DomainPiping => new PipingCollisionCheckTargetCollector( Document, routesInType ),
        _ => throw new InvalidOperationException(),
      } ;
    }

    public override IFailuresPreprocessor CreateFailuresPreprocessor() => new RoutingFailuresPreprocessor( this ) ;
  }
}