using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class PointOnRouteFilters
  {
    public static bool RepresentativeElement( Element elm )
    {
      return elm.GetRepresentativeRouteName() is not { } representativeRouteName || representativeRouteName == elm.GetRouteName() ;
    }
  }
}