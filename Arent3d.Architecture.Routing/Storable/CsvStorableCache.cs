using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.Storable
{
  public class CsvStorableCache : StorableCache<CsvStorableCache, CsvStorable>
  {
    public CsvStorableCache( Document document ) : base( document )
    {
    }

    protected override CsvStorable CreateNewStorable( Document document, string name ) => new CsvStorable( document ) ;
  }
}