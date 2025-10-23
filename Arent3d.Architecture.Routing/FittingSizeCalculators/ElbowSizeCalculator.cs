using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators
{
  internal abstract class ElbowSizeCalculator : SizeCalculatorBase
  {
    private readonly double _diameter ;

    protected ElbowSizeCalculator( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter ) : base( document, mepCurveGenerator, GetStraightLineLength( diameter ) )
    {
      _diameter = diameter ;

      CalculateConnectorPositions() ;
    }

    private static double GetStraightLineLength( double diameter ) => Math.Max( diameter * 50, 1 ) ; // diameter * 50 or 1ft (greater)

    protected override void GenerateFittingFromConnectors( IReadOnlyList<Connector> connectors )
    {
      if ( 2 != connectors.Count ) return ;

      SetDiameter( connectors[ 0 ], _diameter ) ;
      SetDiameter( connectors[ 1 ], _diameter ) ;

      Document.Create.NewElbowFitting( connectors[ 0 ], connectors[ 1 ] ) ;
    }

    private double? _elbowSize ;
    public double ElbowSize => _elbowSize ??= GetElbowSize( ConnectorPositions ) ;

    private static double GetElbowSize( IReadOnlyList<XYZ>? connectorPositions )
    {
      if ( null == connectorPositions || 2 != connectorPositions.Count ) return 0 ;

      return Math.Max( connectorPositions[ 0 ].GetLength(), connectorPositions[ 1 ].GetLength() ) ;
    }
  }

  internal class Elbow45SizeCalculator : ElbowSizeCalculator
  {
    public Elbow45SizeCalculator( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter ) : base( document, mepCurveGenerator, diameter )
    {
    }

    protected override IReadOnlyList<XYZ> EndDirections => new[] { new XYZ( -1, 0, 0 ), new XYZ( 1, 1, 0 ).Normalize() } ;
  }

  internal class Elbow90SizeCalculator : ElbowSizeCalculator
  {
    public Elbow90SizeCalculator( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter ) : base( document, mepCurveGenerator, diameter )
    {
    }

    protected override IReadOnlyList<XYZ> EndDirections => new[] { new XYZ( -1, 0, 0 ), new XYZ( 0, 1, 0 ) } ;
  }
}