using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class RackNotationStorableCache : StorableCache<RackNotationStorableCache, RackNotationStorable>
  {
    public RackNotationStorableCache( Document document ) : base( document )
    {
    }

    protected override RackNotationStorable CreateNewStorable( Document document, string name ) => new RackNotationStorable( document ) ;
  }
}