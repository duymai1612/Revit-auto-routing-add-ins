using System ;
using System.Collections.Generic ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public static class EnumerableExtensions
  {
    public static int FindIndexByVertexTolerance( this IEnumerable<double> source, double value, Document doc )
    {
      var tole = doc.Application.VertexTolerance ;
      return source.FindIndex( d => Math.Abs( d - value ) <= tole ) ;
    }
  }
}