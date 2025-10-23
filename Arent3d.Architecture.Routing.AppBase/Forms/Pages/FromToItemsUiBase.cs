using System.Collections.Generic ;
using System.Windows.Media.Imaging ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public abstract class FromToItemsUiBase
  {
    public IReadOnlyDictionary<string, BitmapImage> FromToTreeIcons { get ; }
    public string TabTitle { get ; }

    protected FromToItemsUiBase( string tabTitle, IReadOnlyDictionary<string, BitmapImage> fromToTreeIcons )
    {
      TabTitle = tabTitle ;
      FromToTreeIcons = fromToTreeIcons ;
    }

    public abstract bool UseHierarchies { get ; }
    public abstract bool ShowSubRoutes { get ; }
  }
}