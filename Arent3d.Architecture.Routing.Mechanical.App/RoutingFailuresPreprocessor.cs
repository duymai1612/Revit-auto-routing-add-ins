using Arent3d.Architecture.Routing.AppBase ;

namespace Arent3d.Architecture.Routing.Mechanical.App
{
  public class RoutingFailuresPreprocessor : RoutingFailuresPreprocessorBase
  {
    public RoutingFailuresPreprocessor( RoutingExecutor executor ) : base( executor )
    {
    }

    protected override bool PreprocessFailureMessage( FailureMessageHandler handler )
    {
      return false ;
    }
  }
}