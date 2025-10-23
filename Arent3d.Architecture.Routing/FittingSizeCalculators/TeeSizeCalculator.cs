using System ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators
{
  internal class TeeSizeCalculator : SizeCalculatorBase
  {
    private readonly double _diameter1 ;
    private readonly double _diameter2 ;

    public TeeSizeCalculator( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter1, double diameter2 ) : base( document, mepCurveGenerator, GetStraightLineLength( diameter1, diameter2 ) )
    {
      _diameter1 = diameter1 ;
      _diameter2 = diameter2 ;

      CalculateConnectorPositions() ;
    }

    private static double GetStraightLineLength( double diameter1, double diameter2 ) => Math.Max( Math.Max( diameter1, diameter2 ) * 50, 1 ) ; // diameter * 50 or 1ft (greater)

    protected override IReadOnlyList<XYZ> EndDirections => new[] { new XYZ( -1, 0, 0 ), new XYZ( 1, 0, 0 ), new XYZ( 0, 1, 0 ) } ;

    protected override void GenerateFittingFromConnectors( IReadOnlyList<Connector> connectors )
    {
      if ( 3 != connectors.Count ) return ;

      SetDiameter( connectors[ 0 ], _diameter1 ) ;
      SetDiameter( connectors[ 1 ], _diameter2 ) ;
      SetDiameter( connectors[ 2 ], _diameter1 ) ;

      Document.Create.NewTeeFitting( connectors[ 0 ], connectors[ 1 ], connectors[ 2 ] ) ;
    }

    private (double HeaderSize, double BranchSize)? _teeSizes ;

    private static (double HeaderSize, double BranchSize) GetTeeSize( IReadOnlyList<XYZ>? connectorPositions )
    {
      if ( null == connectorPositions || 3 != connectorPositions.Count ) return ( 0, 0 ) ;

      var headerSize = Math.Max( connectorPositions[ 0 ].GetLength(), connectorPositions[ 1 ].GetLength() ) ;
      var branchSize = connectorPositions[ 2 ].GetLength() ;
      return ( headerSize, branchSize ) ;
    }

    public double HeaderSize => ( _teeSizes ??= GetTeeSize( ConnectorPositions ) ).HeaderSize ;
    public double BranchSize => ( _teeSizes ??= GetTeeSize( ConnectorPositions ) ).BranchSize ;
  }
}