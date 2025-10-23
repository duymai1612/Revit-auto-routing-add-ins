using System.Collections.Generic ;
using System.Collections.ObjectModel ;
using System.Linq ;
using Arent3d.Architecture.Routing.AppBase.Forms ;
using Autodesk.Revit.ApplicationServices ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Mechanical ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase.Model
{
  public class FromToModel
  {
    private UIDocument UiDoc { get ; set ; }
    private Document Doc { get ; set ; }

    public FromToModel( UIApplication uiApp )
    {
      UiDoc = uiApp.ActiveUIDocument ;
      Doc = UiDoc.Document ;
    }

    /// <summary>
    /// return Hierarchical FromToData for TreeView
    /// </summary>
    /// <returns></returns>
    public IReadOnlyCollection<FromToItem> GetFromToData( AddInType addInType, FromToItemsUiBase fromToItemsUiBase )
    {
      var allRoutes = UiDoc.Document.CollectRoutes( addInType ).ToList() ;

      var fromToItems = FromToItem.CreateRouteFromToItems( Doc, UiDoc, allRoutes, fromToItemsUiBase ) ;

      return fromToItems.ToList() ;
    }
  }
}