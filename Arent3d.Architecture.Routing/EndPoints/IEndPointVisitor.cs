namespace Arent3d.Architecture.Routing.EndPoints
{
  public interface IEndPointVisitor
  {
    void Visit( ConnectorEndPoint endPoint ) ;
    void Visit( RouteEndPoint endPoint ) ;
    void Visit( PassPointEndPoint endPoint ) ;
    void Visit( PassPointBranchEndPoint endPoint ) ;
    void Visit( TerminatePointEndPoint endPoint ) ;
  }

  public interface IEndPointVisitor<out T>
  {
    T Visit( ConnectorEndPoint endPoint ) ;
    T Visit( RouteEndPoint endPoint ) ;
    T Visit( PassPointEndPoint endPoint ) ;
    T Visit( PassPointBranchEndPoint endPoint ) ;
    T Visit( TerminatePointEndPoint endPoint ) ;
  }
}