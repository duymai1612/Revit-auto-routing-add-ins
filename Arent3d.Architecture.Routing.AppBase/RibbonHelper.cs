using System ;
using Arent3d.Revit.I18n ;
using Arent3d.Revit.UI ;
using Autodesk.Revit.DB ;
using Autodesk.Windows ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class RibbonHelper
  {
    
    /// <summary>
    /// Toggle ShowFromToTreeCommandButton depending on FromToTreePanel visible state
    /// </summary>
    /// <param name="isShown"></param>
    public static void ToggleShowFromToTreeCommandButton( bool isShown )
    {
      if ( UIHelper.GetRibbonTabFromName( "App.Routing.TabName".GetAppStringByKey() ) is not { } targetTab ) return ;
      if ( UIHelper.GetRibbonPanelFromName( "App.Panels.Routing.Routing".GetAppStringByKeyOrDefault( "From-To" ), targetTab ) is not { } targetPanel ) return ;
      if ( UIHelper.GetRibbonButtonFromName( "show_from_to_tree_command", targetPanel ) is not { } targetButton ) return ;
      targetButton.LargeImage = UIExtensions.GetImageFromName( isShown ? "MEP.ico" : "PickFrom-To.png" ) ;
    }
  }
}