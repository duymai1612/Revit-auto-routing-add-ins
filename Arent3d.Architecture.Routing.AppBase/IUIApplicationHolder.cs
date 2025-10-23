using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public interface IUIApplicationHolder
  {
    UIApplication? UiApp { get ; }
  }
}