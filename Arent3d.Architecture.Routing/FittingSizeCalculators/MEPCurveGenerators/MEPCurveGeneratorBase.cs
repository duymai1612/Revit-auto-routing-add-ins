using System.Linq ;
using Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public abstract class MEPCurveGeneratorBase : IMEPCurveGenerator
  {
    protected MEPCurveGeneratorBase( Document document, MEPSystemType? systemType, MEPCurveType curveType )
    {
      Document = document ;
      SystemType = systemType ;
      CurveType = curveType ;
    }

    public Document Document { get ; }

    protected MEPSystemType? SystemType { get ; }
    protected MEPCurveType CurveType { get ; }

    public ElementId MEPSystemTypeId => SystemType.GetValidId() ;
    public ElementId MEPCurveTypeId => CurveType.Id ;

    private ElementId? _levelElementId ;

    public abstract MEPCurve? GenerateCurve( XYZ from, XYZ to ) ;

    protected ElementId GetLevelId() => _levelElementId ??= GetDefaultLevelId( Document ).Id ;

    private static Level GetDefaultLevelId( Document document )
    {
      return document.GetAllElements<Level>().FirstOrDefault() ?? Level.Create( document, 0 ) ;
    }
  }
}