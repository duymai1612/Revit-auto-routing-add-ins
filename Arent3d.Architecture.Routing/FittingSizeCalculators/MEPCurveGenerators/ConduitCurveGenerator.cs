using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public class ConduitCurveGenerator : MEPCurveGeneratorBase
  {
    public ConduitCurveGenerator( Document document, MEPSystemType? systemType, MEPCurveType curveType ) : base( document, systemType, curveType )
    {
    }

    public override MEPCurve? GenerateCurve( XYZ from, XYZ to )
    {
      return Conduit.Create( Document, MEPCurveTypeId, from, to, GetLevelId() ) ;
    }
  }
}