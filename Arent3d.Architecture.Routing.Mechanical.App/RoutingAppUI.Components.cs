using Arent3d.Architecture.Routing.Mechanical.App.Commands ;
using Arent3d.Architecture.Routing.Mechanical.App.Commands.Initialization ;
using Arent3d.Architecture.Routing.Mechanical.App.Commands.Routing ;
using Arent3d.Revit.UI.Attributes ;

namespace Arent3d.Architecture.Routing.Mechanical.App
{
  partial class RoutingAppUI
  {
    [Tab( "Mechanical.App.Routing.TabName", VisibilityMode = TabVisibilityMode.NormalDocument )]
    private static class RoutingTab
    {
      [Panel( "arent3d.architecture.routing.init", TitleKey = "Mechanical.App.Panels.Routing.Initialize" )]
      private static class InitPanel
      {
        [Button( typeof( InitializeCommand ), InitializeButton = true )]
        private static class InitializeCommandButton { }

        [Button( typeof( ShowRoutingViewsCommand ), OnlyInitialized = true )]
        private static class ShowRoutingViewsCommandButton { }
      }

      [Panel("arent3d.architecture.routing.routing", TitleKey = "Mechanical.App.Panels.Routing.Routing" )]
      private static class RoutingPanel
      {
        [Button( typeof( PickRoutingCommand ), OnlyInitialized = true )]
        private static class PickRoutingCommandButton { }

        [Button( typeof( AutoRoutingSupplyAirGrillCommand ), OnlyInitialized = true )]
        private static class AutoRoutingSupplyAirGrillCommandButton { }
      }

      [Panel( "arent3d.architecture.routing.monitor", TitleKey = "Mechanical.App.Panels.Routing.Monitor" )]
      private static class MonitorPanel
      {
        [Button( typeof( MonitorSelectionCommand ), AvailabilityType = typeof( Commands.Enabler.MonitorSelectionCommandEnabler ) )]
        private static class MonitorSelectionCommandButton { }
      }

      [Panel( "arent3d.architecture.rc.debug", TitleKey = "App.Panels.Rc.Debug" )]
      private static class DebugPanel
      {

        [Button( typeof( UninitializeCommand ), OnlyInitialized = true )]
        private static class UnInitializeCommandButton
        {
        }
      }
    }
  }
}