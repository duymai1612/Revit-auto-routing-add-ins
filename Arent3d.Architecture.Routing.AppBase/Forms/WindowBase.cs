using System.Windows ;
using System.Windows.Interop ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public class WindowBase : Window
  {
    public WindowBase( UIDocument uiDoc )
    {
      //Set RevitWindow To owner
      var helper = new System.Windows.Interop.WindowInteropHelper( this ) ;
      helper.Owner = uiDoc.Application.MainWindowHandle ;
    }

    public delegate void ClickEventHandler( object sender, RoutedEventArgs e ) ;
  }
}