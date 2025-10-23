using System ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;
using Autodesk.Revit.DB.Mechanical ;
using Autodesk.Revit.DB.Plumbing ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators
{
  public interface IMEPCurveGenerator
  {
    Document Document { get ; }
    ElementId MEPSystemTypeId { get ; }
    ElementId MEPCurveTypeId { get ; }

    MEPCurve? GenerateCurve( XYZ from, XYZ to ) ;
  }

  internal static class MEPCurveGenerator
  {
    public static IMEPCurveGenerator Create( MEPSystemType? mepSystemType, MEPCurveType curveType )
    {
      var document = curveType.Document ;

      return curveType switch
      {
        CableTrayType => new CableTrayCurveGenerator( document, mepSystemType, curveType ),
        ConduitType => new ConduitCurveGenerator( document, mepSystemType, curveType ),
        DuctType => new DuctCurveGenerator( document, mepSystemType, curveType ),
        FlexDuctType => new FlexDuctCurveGenerator( document, mepSystemType, curveType ),
        FlexPipeType => new FlexPipeCurveGenerator( document, mepSystemType, curveType ),
        PipeType => new PipeCurveGenerator( document, mepSystemType, curveType ),
        _ => throw new InvalidOperationException(),
      } ;
    }
  }
}