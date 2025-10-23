using System ;
using Arent3d.Routing ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Defines converters between Revit data structures and auto routing data structures.
  /// </summary>
  /// <remarks>
  /// TODO: Revit Unit (feet) to Arent Unit (m) conversion.
  /// </remarks>
  public static class InterconversionExtensions
  {
    public static Vector3d To3dRaw( this XYZ xyz ) => new( xyz.X, xyz.Y, xyz.Z ) ;
    public static Box3d To3dRaw( this BoundingBoxXYZ xyz ) => new( xyz.Min.To3dRaw(), xyz.Max.To3dRaw() ) ;

    public static XYZ ToXYZRaw( this Vector3d vec ) => new( vec.x, vec.y, vec.z ) ;

    public static IPipeDiameter DiameterValueToPipeDiameter( this double diameter )
    {
      return new PipeDiameter( diameter ) ;
    }

    public static double GetDiameter( this Connector connector )
    {
      return ( connector.Shape switch
      {
        ConnectorProfileType.Oval => connector.Radius * 2,
        ConnectorProfileType.Rectangular => Math.Max( connector.Width, connector.Height ),
        ConnectorProfileType.Round => connector.Radius * 2,
        _ => 0,
      } ) ;
    }

    public static void SetDiameter( this Connector connector, double diameter )
    {
      switch ( connector.Shape ) {
        case ConnectorProfileType.Oval :
          connector.Radius = diameter * 0.5 ;
          break ;

        case ConnectorProfileType.Rectangular :
        {
          var ratio = diameter / Math.Max( connector.Width, connector.Height ) ;
          connector.Width *= ratio ;
          connector.Height *= ratio ;
          break ;
        }

        case ConnectorProfileType.Round :
          connector.Radius = diameter * 0.5 ;
          break ;

        default : throw new ArgumentOutOfRangeException() ;
      }
    }

    private class PipeDiameter : IPipeDiameter
    {
      public PipeDiameter( double diameter )
      {
        Outside = diameter ;
        NPSmm = (int) Math.Floor( diameter * 1000 ) ; // provisional
      }

      public double Outside { get ; }
      public int NPSmm { get ; }
    }
  }
}