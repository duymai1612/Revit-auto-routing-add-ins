using System.Collections.Generic ;
using Arent3d.GeometryLib ;
using MathLib ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public class CollisionBox : IGeometryBody
  {
    private readonly IBox _box ;

    public CollisionBox( Box3d box )
    {
      _box = Box.Create( new LocalCodSys3d( box.Center ), box.Extents ) ;
    }
    
    public CollisionBox(IBox box)
    {
      _box = box ;
    }

    public IReadOnlyCollection<IGeometry> GetGeometries()
    {
      return new List<IGeometry> { _box } ;
    }

    public IReadOnlyCollection<IGeometry> GetGlobalGeometries()
    {
      return GetGeometries() ;
    }

    public Box3d GetGlobalGeometryBox()
    {
      return this.ConvertToBox3d() ;
    }
  }
}