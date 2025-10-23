using System.Collections.Generic ;
using Arent3d.Routing ;

namespace Arent3d.Architecture.Routing
{
  internal class AutoRoutingTargetMerger<TAutoRoutingTarget> where TAutoRoutingTarget : class, IAutoRoutingTarget
  {
    private readonly int _slotCount ;
    private readonly List<TAutoRoutingTarget> _targets ;
    private readonly List<IAutoRoutingResult> _results ;

    public bool IsFullfilled => ( _slotCount == _results.Count ) ;
    
    public AutoRoutingTargetMerger( int slotCount )
    {
      _slotCount = slotCount ;
      _targets = new List<TAutoRoutingTarget>( slotCount ) ;
      _results = new List<IAutoRoutingResult>( slotCount ) ;
    }

    public bool Register( TAutoRoutingTarget target, IAutoRoutingResult result )
    {
      if ( IsFullfilled ) return false ;

      _targets.Add( target ) ;
      _results.Add( result ) ;
      return true ;
    }

    public IReadOnlyCollection<TAutoRoutingTarget> GetAutoRoutingTargets() => _targets ;

    public MergedAutoRoutingResult GetAutoRoutingResult()
    {
      return new MergedAutoRoutingResult( _results ) ;
    }
  }
}