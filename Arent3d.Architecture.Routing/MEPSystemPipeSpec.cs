using System ;
using Arent3d.Architecture.Routing.FittingSizeCalculators ;
using Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators ;
using Arent3d.Routing ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public class MEPSystemPipeSpec : IPipeSpec
  {
    public const double MinimumShortCurveLength = ( 1.0 / 120 ) * 2 ; // twice of the minimum curve length.

    private readonly IFittingSizeCalculator _fittingSizeCalculator ;
    private Document Document => RouteMEPSystem.Document ;

    public double AngleTolerance => RouteMEPSystem.AngleTolerance ;
    public double DiameterTolerance => RouteMEPSystem.DiameterTolerance ;

    public RouteMEPSystem RouteMEPSystem { get ; }

    public string? Name => RouteMEPSystem.MEPSystemType?.Name ;

    public MEPSystemPipeSpec( RouteMEPSystem routeMepSystem, IFittingSizeCalculator fittingSizeCalculator )
    {
      RouteMEPSystem = routeMepSystem ;
      _fittingSizeCalculator = fittingSizeCalculator ;
    }



    private SizeTable<double, double>? _90ElbowSize ;
    private SizeTable<double, double>? _45ElbowSize ;
    private SizeTable<(double HeaderDiameter, double BranchDiameter), (double HeaderLength, double BranchLength)>? _teeSizeLength ;
    private SizeTable<(double, double), double>? _reducerLength ;

    public double GetLongElbowSize( IPipeDiameter diameter )
    {
      return ( _90ElbowSize ??= new SizeTable<double, double>( Calculate90ElbowSize ) ).Get( diameter.Outside ) ;
    }

    public double Get45ElbowSize( IPipeDiameter diameter )
    {
      return ( _45ElbowSize ??= new SizeTable<double, double>( Calculate45ElbowSize ) ).Get( diameter.Outside ) ;
    }

    public double GetTeeHeaderLength( IPipeDiameter header, IPipeDiameter branch )
    {
      if ( JunctionType.Tee == RouteMEPSystem.CurveType.PreferredJunctionType ) {
        return ( _teeSizeLength ??= new SizeTable<(double HeaderDiameter, double BranchDiameter), (double HeaderLength, double BranchLength)>( CalculateTeeLengths ) ).Get( ( header.Outside, branch.Outside ) ).HeaderLength ;
      }
      else {
        return branch.Outside * 0.5 ; // provisional
      }
    }

    public double GetTeeBranchLength( IPipeDiameter header, IPipeDiameter branch )
    {
      if ( JunctionType.Tee == RouteMEPSystem.CurveType.PreferredJunctionType ) {
        return ( _teeSizeLength ??= new SizeTable<(double HeaderDiameter, double BranchDiameter), (double HeaderLength, double BranchLength)>( CalculateTeeLengths ) ).Get( ( header.Outside, branch.Outside ) ).BranchLength ;
      }
      else {
        return branch.Outside * 0.5 ; // provisional
      }
    }


    double IPipeSpec.GetConcentricReducerCombinationLength( IPipeDiameter largeDiameter, IPipeDiameter smallDiameter ) => GetReducerLength( largeDiameter, smallDiameter ) ;
    double IPipeSpec.GetEccentricReducerCombinationLength( IPipeDiameter largeDiameter, IPipeDiameter smallDiameter ) => GetReducerLength( largeDiameter, smallDiameter ) ;

    public double GetReducerLength( IPipeDiameter pipe1, IPipeDiameter pipe2 )
    {
      double diameter1 = pipe1.Outside, diameter2 = pipe2.Outside ;
      
      if ( diameter1 <= 0 || diameter2 <= 0 || Math.Abs( diameter1 - diameter2 ) < DiameterTolerance ) return 0 ;

      if ( diameter1 > diameter2 ) {
        ( diameter1, diameter2 ) = ( diameter2, diameter1 ) ;
      }

      return ( _reducerLength ??= new SizeTable<(double, double), double>( CalculateConcentricReducerLength ) ).Get( ( diameter1, diameter2 ) ) ;
    }

    public double GetWeldMinDistance( IPipeDiameter diameter ) => 0 ;


    private IMEPCurveGenerator? _mepCurveGenerator = null ;
    private IMEPCurveGenerator MEPCurveGenerator => _mepCurveGenerator ??= FittingSizeCalculators.MEPCurveGenerators.MEPCurveGenerator.Create( RouteMEPSystem.MEPSystemType, RouteMEPSystem.CurveType ) ; 

    private double Calculate90ElbowSize( double diameter )
    {
      return _fittingSizeCalculator.Calc90ElbowSize( Document, MEPCurveGenerator, diameter ) + MinimumShortCurveLength ;
    }

    private double Calculate45ElbowSize( double diameter )
    {
      return _fittingSizeCalculator.Calc45ElbowSize( Document, MEPCurveGenerator, diameter ) + MinimumShortCurveLength ;
    }

    private (double Header, double Branch) CalculateTeeLengths( ( double HeaderDiameter, double BranchDiameter) value )
    {
      var tuple = _fittingSizeCalculator.CalculateTeeLengths( Document, MEPCurveGenerator, value.HeaderDiameter, value.BranchDiameter ) ;
      return ( tuple.Header + MinimumShortCurveLength, tuple.Branch + MinimumShortCurveLength ) ;
    }

    private double CalculateConcentricReducerLength( (double, double) value )
    {
      return _fittingSizeCalculator.CalculateReducerLength( Document, MEPCurveGenerator, value.Item1, value.Item2 ) + ( MinimumShortCurveLength * 2 ) ;
    }
  }
}