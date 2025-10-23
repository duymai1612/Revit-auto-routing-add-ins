using System ;
using System.Collections.Generic ;
using System.Windows.Media.Imaging ;
using Arent3d.Architecture.Routing.AppBase.Forms ;
using Arent3d.Revit.I18n ;

namespace Arent3d.Architecture.Routing.Mechanical.App.Forms
{
  public class FromToItemsUi : FromToItemsUiBase
  {
    private static readonly IReadOnlyDictionary<string, BitmapImage> FromToTreeIconDictionary = new Dictionary<string, BitmapImage>
    {
      { "RouteItem", new BitmapImage( new Uri( "../resources/DeleteAllFrom-To.png", UriKind.Relative ) ) },
      { "ConnectorItem", new BitmapImage( new Uri( "../resources/DeleteAllPS.png", UriKind.Relative ) ) },
      { "SubRouteItem", new BitmapImage( new Uri( "../resources/DeleteFrom-To.png", UriKind.Relative ) ) },
      { "PassPointItem", new BitmapImage( new Uri( "../resources/ExportFromTo.png", UriKind.Relative ) ) },
      { "TerminatePointItem", new BitmapImage( new Uri( "../resources/ExportPS.png", UriKind.Relative ) ) },
    } ;

    public override bool UseHierarchies => true ;
    public override bool ShowSubRoutes => true ;

    public FromToItemsUi() : base( "Dialog.Forms.FromToTree.Title".GetAppStringByKeyOrDefault( "Mechanical From-To" ), FromToTreeIconDictionary )
    {
    }
  }
}