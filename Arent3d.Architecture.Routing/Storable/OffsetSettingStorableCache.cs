using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class OffsetSettingStorableCache : StorableCache<OffsetSettingStorableCache, OffsetSettingStorable>
  {
    public OffsetSettingStorableCache( Document document ) : base( document )
    {
    }

    protected override OffsetSettingStorable CreateNewStorable( Document document, string name ) => new OffsetSettingStorable( document ) ;
  }
}