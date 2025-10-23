using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Revit.UI ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Commands.Initialization
{
  public abstract class UninitializeCommandBase : ExternalCommandBase
  {
    protected override string GetTransactionName() => "TransactionName.Commands.Rc.UnInitialize".GetAppStringByKeyOrDefault( "Erase all addin data" ) ;

    protected override ExecutionResult Execute( Document document, TransactionWrapper transaction )
    {
      RouteCache.Release( document ) ;
      document.UnsetupRoutingFamiliesAndParameters() ;
      document.DeleteAllDerivedStorables( typeof( AppInfo ).Assembly ) ;

      return ExecutionResult.Succeeded ;
    }
  }
}