using System ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Plumbing ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public class FlexPipeCurveGenerator : MEPCurveGeneratorBase
  {
    public FlexPipeCurveGenerator( Document document, MEPSystemType? systemType, MEPCurveType curveType ) : base( document, systemType, curveType )
    {
    }

    public override MEPCurve? GenerateCurve( XYZ from, XYZ to )
    {
      return FlexPipe.Create( Document, MEPSystemTypeId, MEPCurveTypeId, GetLevelId(), from, to, Array.Empty<XYZ>() ) ;
    }
  }
}