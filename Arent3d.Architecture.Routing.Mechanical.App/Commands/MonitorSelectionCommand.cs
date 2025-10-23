using System.ComponentModel ;
using Autodesk.Revit.Attributes ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.Mechanical.App.Commands
{
  [Transaction( TransactionMode.Manual )]
  [DisplayName( "MonitorSelection" )]
  public class MonitorSelectionCommand : IExternalCommand
  {
    public Result Execute( ExternalCommandData commandData, ref string message, ElementSet elements )
    {
      TaskDialog.Show( "monitor selection", "selected" ) ;

      return Result.Succeeded ;
    }
  }
}