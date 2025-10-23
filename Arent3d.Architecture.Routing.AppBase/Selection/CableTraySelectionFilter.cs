using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;
using Autodesk.Revit.UI.Selection ;

namespace Arent3d.Architecture.Routing.AppBase.Selection
{
  public class CableTraySelectionFilter : ISelectionFilter
  {
    public static ISelectionFilter Instance { get ; } = new CableTraySelectionFilter() ;

    private CableTraySelectionFilter()
    {
    }
    
    public bool AllowElement( Element elem )
    {
      return ( BuiltInCategory.OST_CableTray == elem.GetBuiltInCategory() || BuiltInCategory.OST_CableTrayFitting == elem.GetBuiltInCategory()) && elem is FamilyInstance or CableTray ;
    }

    public bool AllowReference( Reference reference, XYZ position ) => false ;
  }
}
