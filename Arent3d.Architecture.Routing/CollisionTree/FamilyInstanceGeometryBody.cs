using System ;
using System.Collections.Generic ;
using Arent3d.GeometryLib ;
using Arent3d.Routing ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing.CollisionTree
{
  public class FamilyInstanceGeometryBody : IGeometryBody
  {
    private readonly View? _view ;
    private readonly IReadOnlyCollection<GeometryObject> _gObjs ;

    public FamilyInstance FamilyInstance { get ; }
    
    public bool IsStructure { get ; } // TODO

    public FamilyInstanceGeometryBody( FamilyInstance instance )
    {
      FamilyInstance = instance ;

      _view = null ;
      var gElm = FamilyInstance.get_Geometry( new Options { DetailLevel = ViewDetailLevel.Coarse, ComputeReferences = true, IncludeNonVisibleObjects = false, View = _view } ) ;
      _gObjs = gElm.EnumerateAll() ;
    }

    public IReadOnlyCollection<IGeometry> GetGeometries() => GetGlobalGeometries() ;

    public IReadOnlyCollection<IGeometry> GetGlobalGeometries()
    {
      var list = new List<IGeometry>() ;
      foreach ( var gObj in _gObjs ) {
        // TODO
      }

      return list ;
    }

    public Box3d GetGlobalGeometryBox()
    {
      return FamilyInstance.get_BoundingBox( _view ).To3dRaw() ;
    }

    public IRouteCondition? GetRouteCondition()
    {
      // TODO
      return null ;
    }
  }
}