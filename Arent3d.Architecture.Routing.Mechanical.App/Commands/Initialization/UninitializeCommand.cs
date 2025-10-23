using Arent3d.Architecture.Routing.AppBase.Commands.Initialization ;
using Arent3d.Revit.UI ;
using Autodesk.Revit.Attributes ;

namespace Arent3d.Architecture.Routing.Mechanical.App.Commands.Initialization
{
  [Transaction( TransactionMode.Manual )]
  [DisplayNameKey( "Mechanical.App.Commands.Initialization.UnnitializeCommand", DefaultString = "Erase all addin data" )]
  [Image( "resources/Initialize.png", ImageType = ImageType.Large )]
  public class UninitializeCommand : UninitializeCommandBase
  {
  }
}