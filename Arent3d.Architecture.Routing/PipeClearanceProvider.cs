using Arent3d.Routing ;

namespace Arent3d.Architecture.Routing
{
  public class PipeClearanceProvider : IPipeClearanceProvider
  {
    public static PipeClearanceProvider Instance { get ; } = new PipeClearanceProvider() ;

    private PipeClearanceProvider()
    {
    }

    public double Get( IRouteCondition route1, IRouteCondition route2 ) => 0.01 ;

    public double GetClearanceToObstacle( IRouteCondition route ) => 0.05 ; // about 2 in
  }
}