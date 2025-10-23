using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Windows ;
using System.Windows.Controls ;
using System.Windows.Media ;
using Arent3d.Revit.I18n ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Windows ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class UIHelper
  {
    /// <summary>
    /// Get LabelName From CurveType
    /// </summary>
    /// <param name="targetStrings"></param>
    /// <returns></returns>
    public static string GetTypeLabel( string targetStrings )
    {
      if ( targetStrings.EndsWith( "Type" ) ) {
        targetStrings = targetStrings.Substring( 0, targetStrings.Length - 4 ) + " Type" ;
      }

      return targetStrings ;
    }

    public static RibbonTab? GetRibbonTabFromName( string? targetTabName ) => ComponentManager.Ribbon.Tabs.FirstOrDefault( t => t.Id == targetTabName ) ;

    public static RibbonPanel? GetRibbonPanelFromName( string targetPanelName, RibbonTab? targetRibbonTab ) => targetRibbonTab?.Panels.FirstOrDefault( p => p.Source.Title == targetPanelName ) ;

    public static RibbonButton? GetRibbonButtonFromName( string targetButtonCommand, RibbonPanel? targetRibbonPanel )
    {
      var targetItemName = "CustomCtrl_%" + targetRibbonPanel?.Source.Id + "%arent3d.architecture.routing.appbase.commands.routing." + targetButtonCommand ;
      return targetRibbonPanel?.Source.Items.OfType<RibbonButton>().FirstOrDefault( item => item.Id == targetItemName ) ;
    }

    public static int GetPositionAfterButton( string s )
    {
      var index = ComponentManager.QuickAccessToolBar.Items.FindIndex( item => item.Id == s ) ;
      return ( 0 <= index ? index + 1 : -1 ) ;
    }

    public static void PlaceButtonOnQuickAccess( int position, Autodesk.Windows.RibbonItem ribbonItem )
    {
      if ( position < 0 ) {
        ComponentManager.QuickAccessToolBar.AddStandardItem( ribbonItem ) ;
      }
      else {
        ComponentManager.QuickAccessToolBar.InsertStandardItem( position, ribbonItem ) ;
      }
    }

    public static void RemovePanelFromTab( RibbonTab ribbonTab, Autodesk.Windows.RibbonPanel ribbonPanel )
    {
      ribbonTab.Panels.Remove( ribbonPanel ) ;
    }

    public static void RemoveTabFromRibbon( RibbonTab ribbonTab )
    {
      if ( ribbonTab.Panels.Count != 0 ) {
        return ;
      }

      ribbonTab.IsVisible = false ;
    }

    public static int FindClosestIndex( IList<double> list, double value )
    {
      var closest = list.OrderBy(item => Math.Abs(value - item)).First();
      return list.IndexOf(closest);
    }
    
    public static T? FindAncestor<T>(this DependencyObject depObj)
      where T : DependencyObject
    {
      while (depObj != null)
      {
        if (depObj is T target)
        {
          return target;
        }
        depObj = VisualTreeHelper.GetParent(depObj);
      }
      return null;
    }
  }
}