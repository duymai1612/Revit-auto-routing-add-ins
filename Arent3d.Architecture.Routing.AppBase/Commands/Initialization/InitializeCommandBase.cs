using System ;
using Arent3d.Revit.I18n ;
using Arent3d.Revit.UI ;
using Arent3d.Architecture.Routing.AppBase.Commands ;
using Autodesk.Revit.Attributes ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;
using ImageType = Arent3d.Revit.UI.ImageType ;

namespace Arent3d.Architecture.Routing.AppBase.Commands.Initialization
{
  public abstract class InitializeCommandBase : IExternalCommand
  {
    public Result Execute( ExternalCommandData commandData, ref string message, ElementSet elements )
    {
      var document = commandData.Application.ActiveUIDocument.Document ;
      if ( document.RoutingSettingsAreInitialized() ) return Result.Succeeded ;

      try {
        var result = document.Transaction( "TransactionName.Commands.Initialization.Initialize".GetAppStringByKeyOrDefault( "Setup Routing" ), _ =>
        {
          return Setup( document ) ? Result.Succeeded : Result.Failed ;
        } ) ;

        if ( Result.Failed == result ) {
          TaskDialog.Show( "Dialog.Commands.Initialization.Dialog.Title.Error".GetAppStringByKeyOrDefault( null ), "Dialog.Commands.Initialization.Dialog.Body.Error.FailedToSetup".GetAppStringByKeyOrDefault( null ) ) ;
        }

        return result ;
      }
      catch ( Exception e ) {
        CommandUtils.DebugAlertException( e ) ;
        return Result.Failed ;
      }
    }

    protected virtual bool Setup( Document document )
    {
      return document.SetupRoutingFamiliesAndParameters() ;
    }
  }
}