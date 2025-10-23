using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class PickUpStorableCache : StorableCache<PickUpStorableCache, PickUpStorable>
  {
    public PickUpStorableCache( Document document ) : base( document )
    {
    }

    protected override PickUpStorable CreateNewStorable( Document document, string name ) => new PickUpStorable( document ) ;
  }
}