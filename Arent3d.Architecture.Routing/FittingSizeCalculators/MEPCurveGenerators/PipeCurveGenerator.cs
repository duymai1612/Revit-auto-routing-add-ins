using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Plumbing ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public class PipeCurveGenerator : MEPCurveGeneratorBase
  {
    public PipeCurveGenerator( Document document, MEPSystemType? systemType, MEPCurveType curveType ) : base( document, systemType, curveType )
    {
    }

    public override MEPCurve? GenerateCurve( XYZ from, XYZ to )
    {
      return Pipe.Create( Document, MEPSystemTypeId, MEPCurveTypeId, GetLevelId(), from, to ) ;
    }
  }
}