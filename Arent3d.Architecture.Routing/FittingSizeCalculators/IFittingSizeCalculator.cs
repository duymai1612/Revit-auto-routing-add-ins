using Arent3d.Architecture.Routing.FittingSizeCalculators.MEPCurveGenerators ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.FittingSizeCalculators
{
  public interface IFittingSizeCalculator
  {
    double Calc90ElbowSize( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter ) ;
    double Calc45ElbowSize( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter ) ;
    (double Header, double Branch) CalculateTeeLengths( Document document, IMEPCurveGenerator mepCurveGenerator, double headerDiameter, double branchDiameter ) ;
    double CalculateReducerLength( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter1, double diameter2 ) ;
  }

  public class DefaultFittingSizeCalculator : IFittingSizeCalculator
  {
    public static IFittingSizeCalculator Instance { get ; } = new DefaultFittingSizeCalculator() ;

    private DefaultFittingSizeCalculator()
    {
    }

    public double Calc90ElbowSize( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter )
    {
      var calculator = new Elbow90SizeCalculator( document, mepCurveGenerator, diameter ) ;
      return calculator.ElbowSize ;
    }

    public double Calc45ElbowSize( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter )
    {
      var calculator = new Elbow45SizeCalculator( document, mepCurveGenerator, diameter ) ;
      return calculator.ElbowSize ;
    }

    public (double Header, double Branch) CalculateTeeLengths( Document document, IMEPCurveGenerator mepCurveGenerator, double headerDiameter, double branchDiameter )
    {
      var calculator = new TeeSizeCalculator( document, mepCurveGenerator, headerDiameter, branchDiameter ) ;
      return ( calculator.HeaderSize, calculator.BranchSize ) ;
    }

    public double CalculateReducerLength( Document document, IMEPCurveGenerator mepCurveGenerator, double diameter1, double diameter2 )
    {
      var calculator = new ReducerSizeCalculator( document, mepCurveGenerator, diameter1, diameter2 ) ;
      return calculator.Size1 + calculator.Size2 ;
    }
  }
  
}