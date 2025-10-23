using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI.Selection ;

namespace Arent3d.Architecture.Routing.AppBase.Selection
{
  public class ConnectorFamilySelectionFilter : ISelectionFilter
  {
    public static ISelectionFilter Instance { get ; } = new ConnectorFamilySelectionFilter() ;

    private ConnectorFamilySelectionFilter()
    {
    }
    
    public bool AllowElement( Element elem )
    {
      return ( BuiltInCategory.OST_ElectricalFixtures == elem.GetBuiltInCategory() ) && elem is FamilyInstance ;
    }

    public bool AllowReference( Reference reference, XYZ position ) => false ;
  }
}