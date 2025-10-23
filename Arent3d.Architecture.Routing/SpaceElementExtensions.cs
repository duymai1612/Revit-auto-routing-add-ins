using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public enum SpaceType
  {
    Invalid = -1,
    Parent = 0
  }

  public static class SpaceElementExtensions
  {
    public static int GetSpaceBranchNumber( this Element space )
    {
      if ( ! space.HasParameter( BranchNumberParameter.BranchNumber ) ) return (int) SpaceType.Invalid ;
      return space.GetPropertyInt( BranchNumberParameter.BranchNumber ) ;
    }
  }
}