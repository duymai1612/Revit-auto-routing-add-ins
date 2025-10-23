using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class CeedStorableCache : StorableCache<CeedStorableCache, CeedStorable>
  {
    public CeedStorableCache( Document document ) : base( document )
    {
    }

    protected override CeedStorable CreateNewStorable( Document document, string name ) => new CeedStorable( document ) ;
  }
}