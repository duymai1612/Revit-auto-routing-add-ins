using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;
using Autodesk.Revit.UI.Selection ;

namespace Arent3d.Architecture.Routing.AppBase.Selection
{
  public class ConstructionItemSelectionFilter : ISelectionFilter
  {
    public static ISelectionFilter Instance { get ; } = new ConstructionItemSelectionFilter() ;

    private ConstructionItemSelectionFilter()
    {
    }
    
    public bool AllowElement( Element elem )
    {
      return ( BuiltInCategory.OST_Conduit == elem.GetBuiltInCategory() || 
               BuiltInCategory.OST_ConduitFitting == elem.GetBuiltInCategory() ||
               BuiltInCategory.OST_ConduitRun == elem.GetBuiltInCategory() ||
               BuiltInCategory.OST_ElectricalFixtures == elem.GetBuiltInCategory() ||
               BuiltInCategory.OST_CableTray == elem.GetBuiltInCategory() ||
               BuiltInCategory.OST_CableTrayFitting == elem.GetBuiltInCategory()) 
             && elem is FamilyInstance or CableTray or Conduit;
    }

    public bool AllowReference( Reference reference, XYZ position ) => false ;
  }
}
