using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class HeightSettingStorableCache : StorableCache<HeightSettingStorableCache, HeightSettingStorable>
  {
    public HeightSettingStorableCache( Document document ) : base( document )
    {
    }

    protected override HeightSettingStorable CreateNewStorable( Document document, string name ) => new HeightSettingStorable( document ) ;
  }
}