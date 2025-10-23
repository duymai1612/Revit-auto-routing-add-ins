using Arent3d.Architecture.Routing.AppBase ;
using Arent3d.Revit.I18n ;
using Autodesk.Revit.DB.Events ;
using Autodesk.Windows ;
using RibbonTab = Autodesk.Windows.RibbonTab ;

namespace Arent3d.Architecture.Routing.Mechanical.App
{
  internal static class MonitorSelectionApplicationEvent
  {
    public static void MonitorSelectionApplicationInitialized( object sender, ApplicationInitializedEventArgs e )
    {
      RibbonTab? selectionTab = null ;
      RibbonPanel? selectionPanel = null ;
      RibbonItem? selectionButton = null ;
      string? targetTabName = "Mechanical.App.Routing.TabName".GetAppStringByKey() ;


      selectionTab = UIHelper.GetRibbonTabFromName( targetTabName ) ;
      if ( selectionTab == null ) {
        return ;
      }
      else {
        foreach ( var panel in selectionTab.Panels ) {
          if ( panel.Source.Title == "Mechanical.App.Panels.Routing.Monitor".GetAppStringByKeyOrDefault("Monitor Selection") ) {
            selectionPanel = panel ;
            foreach ( var item in panel.Source.Items ) {
              if ( item.Id == "CustomCtrl_%CustomCtrl_%" + targetTabName + "%arent3d.architecture.routing.monitor%arent3d.architecture.routing.mechanical.app.commands.monitor_selection_command" ) {
                selectionButton = item ;
                break ;
              }
              else {
                continue ;
              }
            }
          }
          else {
            continue;
          }
        }
      }

      if ( selectionPanel != null && selectionButton != null ) {
        var position = UIHelper.GetPositionAfterButton( "ID_REVIT_FILE_PRINT" ) ;

        UIHelper.PlaceButtonOnQuickAccess( position, selectionButton ) ;
        // Remove Panel
        UIHelper.RemovePanelFromTab(selectionTab, selectionPanel);
      }
    }
  }
}