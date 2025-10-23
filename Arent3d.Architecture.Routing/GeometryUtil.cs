using System ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public static class GeometryUtil
  {
    public static bool IsPerpendicularTo( this XYZ dir1, XYZ dir2, double angleTolerance )
    {
      return Math.Abs( Math.PI / 2 - dir1.AngleTo( dir2 ) ) < angleTolerance ;
    }
  }
}