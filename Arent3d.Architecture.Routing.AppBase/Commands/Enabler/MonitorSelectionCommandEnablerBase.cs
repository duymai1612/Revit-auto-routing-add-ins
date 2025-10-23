using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.AppBase.ViewModel ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase.Commands.Enabler
{
  public abstract class MonitorSelectionCommandEnablerBase : IExternalCommandAvailability
  {
    private ElementId? _previousSelectedRouteElementId = null ;

    protected abstract AddInType GetAddInType() ;

    public bool IsCommandAvailable( UIApplication uiApp, CategorySet selectedCategories )
    {
      var uiDoc = uiApp.ActiveUIDocument ;

      //If no Doc
      if ( uiDoc == null ) {
        return false ;
      }

      // Raise the SelectionChangedEvent
      var selectedRoutes = PointOnRoutePicker.PickedRoutesFromSelections( uiDoc ).EnumerateAll() ;

      ElementId? selectedElementId = null ;


      var routeNameLst = new List<string>() ;
      var elementIds = uiDoc.Selection.GetElementIds() ;

      foreach ( ElementId eid in elementIds ) {
        var elem = uiDoc.Document.GetElement( eid ) ;
        var routeName = elem.GetRouteName() ;
        if ( routeName is null ) continue ;
        if ( routeNameLst.Contains( routeName ) == false ) {
          routeNameLst.Add( routeName ) ;
        }
      }

      if ( routeNameLst.Count > 1 ) {
        _previousSelectedRouteElementId = null ;
      }

      // if route selected
      if ( selectedRoutes.FirstOrDefault() is { } selectedRoute ) {
        selectedElementId = selectedRoute.OwnerElement?.Id ;

        _previousSelectedRouteElementId = selectedElementId ;
      }

      // if Connector selected
      else if ( uiDoc.Document.CollectRoutes( GetAddInType() ).SelectMany( r => r.GetAllConnectors() ).Any( c => uiDoc.Selection.GetElementIds().Contains( c.Owner.Id ) ) ) {
        selectedElementId = uiDoc.Selection.GetElementIds().FirstOrDefault() ;
        _previousSelectedRouteElementId = selectedElementId ;
      }

      else if ( _previousSelectedRouteElementId != null ) {
        _previousSelectedRouteElementId = null ;
      }


      return false ;
    }
  }
}