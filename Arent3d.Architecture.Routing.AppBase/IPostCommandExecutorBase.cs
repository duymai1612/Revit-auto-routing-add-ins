using System.Collections.Generic ;
using Arent3d.Architecture.Routing.AppBase.Forms ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public interface IPostCommandExecutorBase
  {
    void ChangeRouteNameCommand( Route route, string newName ) ;

    void ApplySelectedFromToChangesCommand( Route route, IReadOnlyCollection<SubRoute> subRoutes, RouteProperties properties ) ;
  }
}