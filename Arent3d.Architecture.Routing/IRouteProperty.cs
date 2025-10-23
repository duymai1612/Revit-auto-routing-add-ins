using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public interface IRouteProperty
  {
    MEPSystemType? GetSystemType() ;
    MEPCurveType GetCurveType() ;
    double GetDiameter() ;
    bool GetRouteOnPipeSpace() ;
    FixedHeight? GetFromFixedHeight() ;
    FixedHeight? GetToFixedHeight() ;
    AvoidType GetAvoidType() ;
    Opening? GetShaft() ;
  }
}