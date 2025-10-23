using Arent3d.Routing ;
using Arent3d.Routing.Conditions ;

namespace Arent3d.Architecture.Routing
{
  public class MEPSystemRouteCondition : IRouteCondition
  {
    private const string DefaultFluidPhase = "None" ;

    public IPipeDiameter Diameter { get ; }
    public double DiameterPipeAndInsulation => Diameter.Outside ;
    public double DiameterFlangeAndInsulation => Diameter.Outside ; // provisional
    IPipeSpec IRouteCondition.Spec => Spec ;
    public MEPSystemPipeSpec Spec { get ; }
    public ProcessConstraint ProcessConstraint { get ; }
    public string FluidPhase => DefaultFluidPhase ;

    public MEPSystemRouteCondition( MEPSystemPipeSpec pipeSpec, double diameter, AvoidType avoidType )
    {
      Spec = pipeSpec ;
      Diameter = diameter.DiameterValueToPipeDiameter() ;
      ProcessConstraint = (ProcessConstraint) avoidType ;
    }
  }
}