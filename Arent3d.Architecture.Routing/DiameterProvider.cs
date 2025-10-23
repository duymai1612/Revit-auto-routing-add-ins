using Arent3d.Routing ;

namespace Arent3d.Architecture.Routing
{
  public class DiameterProvider : IDiameterProvider
  {
    public static DiameterProvider Instance { get ; } = new DiameterProvider() ;

    private DiameterProvider()
    {
    }

    public IPipeDiameter Get( double r ) => r.DiameterValueToPipeDiameter() ;
  }
}