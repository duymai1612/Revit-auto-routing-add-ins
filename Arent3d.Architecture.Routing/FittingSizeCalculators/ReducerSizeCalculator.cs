using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators
{
  internal class ReducerSizeCalculator : SizeCalculatorBase
  {
    private readonly double _diameter1 ;
    private readonly double _diameter2 ;

    public ReducerSizeCalculator( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter1, double diameter2 ) : base( document, mepCurveGenerator, GetStraightLineLength( diameter1, diameter2 ) )
    {
      _diameter1 = diameter1 ;
      _diameter2 = diameter2 ;

      CalculateConnectorPositions() ;
    }

    private static double GetStraightLineLength( double diameter1, double diameter2 ) => Math.Max( Math.Max( diameter1, diameter2 ) * 50, 1 ) ; // diameter * 50 or 1ft (greater)

    protected override IReadOnlyList<XYZ> EndDirections => new[] { new XYZ( -1, 0, 0 ), new XYZ( 1, 0, 0 ) } ;

    protected override void GenerateFittingFromConnectors( IReadOnlyList<Connector> connectors )
    {
      if ( 2 != connectors.Count ) return ;

      SetDiameter( connectors[ 0 ], _diameter1 ) ;
      SetDiameter( connectors[ 1 ], _diameter2 ) ;

      Document.Create.NewTransitionFitting( connectors[ 0 ], connectors[ 1 ] ) ;
    }

    private (double Size1, double Size2)? _reducerSizes ;

    private static (double Size1, double Size2) GetReducerSize( IReadOnlyList<XYZ>? connectorPositions )
    {
      if ( null == connectorPositions || 2 != connectorPositions.Count ) return ( 0, 0 ) ;

      return ( connectorPositions[ 0 ].GetLength(), connectorPositions[ 1 ].GetLength() ) ;
    }

    public double Size1 => ( _reducerSizes ??= GetReducerSize( ConnectorPositions ) ).Size1 ;
    public double Size2 => ( _reducerSizes ??= GetReducerSize( ConnectorPositions ) ).Size2 ;
  }
}