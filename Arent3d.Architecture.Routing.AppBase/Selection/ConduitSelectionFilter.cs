using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI.Selection ;
using Autodesk.Revit.DB.Electrical ;

namespace Arent3d.Architecture.Routing.AppBase.Selection
{
  public class ConduitSelectionFilter : ISelectionFilter
  {
    public static ISelectionFilter Instance { get ; } = new ConduitSelectionFilter() ;

    private ConduitSelectionFilter()
    {
    }

    public bool AllowElement( Element elem )
    {
      return ( BuiltInCategorySets.Conduits.Any( p => p == elem.GetBuiltInCategory() )
               && elem is FamilyInstance or Conduit ) ;
    }

    public bool AllowReference( Reference reference, XYZ position ) => false ;
  }
}