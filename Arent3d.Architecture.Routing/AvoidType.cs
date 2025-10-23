using Arent3d.Routing.Conditions ;


namespace Arent3d.Architecture.Routing
{
  public enum AvoidType
  {
    Whichever = (int) ProcessConstraint.None,
    NoAvoid = (int) ProcessConstraint.NoPocket,
    AvoidAbove = (int) ProcessConstraint.NoDrainPocket,
    AvoidBelow = (int) ProcessConstraint.NoVentPocket,
  }
}