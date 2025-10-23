using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable.StorableConverter
{
  public class DetailSymbolStorableCache : StorableCache<DetailSymbolStorableCache, DetailSymbolStorable>
  {
    public DetailSymbolStorableCache( Document document ) : base( document )
    {
    }

    protected override DetailSymbolStorable CreateNewStorable( Document document, string name ) => new DetailSymbolStorable( document ) ;
  }
}