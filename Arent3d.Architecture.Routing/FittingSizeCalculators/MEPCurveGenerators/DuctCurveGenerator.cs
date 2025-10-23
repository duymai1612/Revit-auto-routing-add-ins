using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Mechanical ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public class DuctCurveGenerator : MEPCurveGeneratorBase
  {
    public DuctCurveGenerator( Document document, MEPSystemType? systemType, MEPCurveType curveType ) : base( document, systemType, curveType )
    {
    }

    public override MEPCurve? GenerateCurve( XYZ from, XYZ to )
    {
      return Duct.Create( Document, MEPSystemTypeId, MEPCurveTypeId, GetLevelId(), from, to ) ;
    }
  }
}