using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class CnsSettingStorableCache : StorableCache<CnsSettingStorableCache, CnsSettingStorable>
  {
    public CnsSettingStorableCache( Document document ) : base( document )
    {
    }

    protected override CnsSettingStorable CreateNewStorable( Document document, string name ) => new CnsSettingStorable( document ) ;
  }
}
