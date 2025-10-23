using Arent3d.Routing ;

namespace Arent3d.Architecture.Routing
{
  public interface IPseudoTerminalPoint : IRouteVertex
  {
    IRouteVertex BaseRouteVertex { get ; }
  }
}