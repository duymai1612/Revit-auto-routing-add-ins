using System ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Mechanical ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public class FlexDuctCurveGenerator : MEPCurveGeneratorBase
  {
    public FlexDuctCurveGenerator( Document document, MEPSystemType? systemType, MEPCurveType curveType ) : base( document, systemType, curveType )
    {
    }

    public override MEPCurve? GenerateCurve( XYZ from, XYZ to )
    {
      return FlexDuct.Create( Document, MEPSystemTypeId, MEPCurveTypeId, GetLevelId(), from, to, Array.Empty<XYZ>() ) ;
    }
  }
}