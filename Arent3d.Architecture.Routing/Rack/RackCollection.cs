using System ;
using System.Collections ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Routing ;
using Arent3d.Utility ;

namespace Arent3d.Architecture.Routing.Rack
{
  public class RackCollection : IStructureGraph
  {
    private readonly HashSet<Rack> _racks = new() ;
    private readonly HashSet<(ILayerStack, ILayerStack)> _links = new() ;

    public bool AddRack( Rack rack )
    {
      return _racks.Add( rack ) ;
    }

    public bool RemoveRack( Rack rack )
    {
      if ( false == _racks.Remove( rack ) ) return false ;

      _links.RemoveWhere( tuple => ( tuple.Item1 == rack || tuple.Item2 == rack ) ) ;
      
      return true ;
    }

    public bool AddLink( Rack rack1, Rack rack2 )
    {
      if ( ! _racks.Contains( rack1 ) || ! _racks.Contains( rack2 ) ) return false ;

      if ( _links.Contains( ( (ILayerStack) rack2, (ILayerStack) rack1 ) ) ) return false ;

      return _links.Add( ( (ILayerStack) rack1, (ILayerStack) rack2 ) ) ;
    }

    public bool RemoveLink( Rack rack1, Rack rack2 )
    {
      return _links.Remove( ( (ILayerStack) rack2, (ILayerStack) rack1 ) ) || _links.Remove( ( (ILayerStack) rack1, (ILayerStack) rack2 ) ) ;
    }

    public void Clear()
    {
      _racks.Clear() ;
      _links.Clear() ;
    }

    public int RackCount => _racks.Count ;
    public int LinkCount => _links.Count ;
    
    public IEnumerable<ILayerStack> Nodes => _racks ;

    public IEnumerable<(ILayerStack, ILayerStack)> Edges => _links ;
    
    public void CreateLinkages()
    {
      _links.Clear() ;

      _links.UnionWith( FindLinkages( _racks ) ) ;
    }

    private static IList<(ILayerStack, ILayerStack)> FindLinkages( IReadOnlyCollection<Rack> racks )
    {
      if ( racks.Count == 0 || racks.Count == 1 ) return Array.Empty<(ILayerStack, ILayerStack)>() ;

      var mainPairs = new Dictionary<Rack, IList<Rack>>() ;
      var auxPairs = new Dictionary<Rack, IList<Rack>>() ;

      void Register( Rack r1, Rack r2 )
      {
        Add( mainPairs, r1, r2 ) ;
        Add( auxPairs, r2, r1 ) ;
      }

      bool HasLink( Rack r ) => mainPairs.ContainsKey( r ) || auxPairs.ContainsKey( r ) ;

      var array = racks.ToArray() ;
      for ( int i = 0 ; i < array.Length ; ++i ) {
        var rack1 = array[ i ] ;
        var looserPartners = new List<Rack>() ;

        for ( int j = i + 1 ; j < array.Length ; ++j ) {
          var rack2 = array[ j ] ;

          if ( rack1.IsExactIntersect( rack2 ) ) {
            Register( rack1, rack2 ) ;
          }
          else if ( rack1.IsLooseIntersect( rack2 ) ) {
            looserPartners.Add( rack2 ) ;
          }
        }

        if ( HasLink( rack1 ) && 0 == looserPartners.Count ) continue ;

        foreach ( var rack2 in looserPartners ) {
          Register( rack1, rack2 ) ;
        }
      }

      return mainPairs.SelectMany( item => item.Value.Select( r => ( (ILayerStack) item.Key, (ILayerStack) r ) ) ).ToList() ;
    }

    private static void Add( IDictionary<Rack, IList<Rack>> pairs, Rack rack1, Rack rack2 )
    {
      if ( pairs.TryGetValue( rack1, out var list ) ) {
        list.Add( rack2 ) ;
      }
      else {
        pairs.Add( rack1, new List<Rack> { rack2 } ) ;
      }
    }
  }
}