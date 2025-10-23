using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Threading ;
using Arent3d.Routing ;
using Arent3d.Utility ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Base class of route generators.
  /// </summary>
  public abstract class RouteGeneratorBase<TAutoRoutingTarget> where TAutoRoutingTarget : class, IAutoRoutingTarget
  {
    /// <summary>
    /// When overridden in a derived class, returns routing targets to generate routes.
    /// </summary>
    protected abstract IReadOnlyList<IReadOnlyCollection<TAutoRoutingTarget>> RoutingTargetGroups { get ; }

    /// <summary>
    /// When overridden in a derived class, returns an collision check tree.
    /// </summary>
    protected abstract CollisionTree.CollisionTree CollisionCheckTree { get ; }

    /// <summary>
    /// When overridden in a derived class, returns a structure graph.
    /// </summary>
    protected abstract IStructureGraph StructureGraph { get ; }
    
    /// <summary>
    /// When overridden in a derived class, this method is called before all route generations. It is good to preprocess for an execution.
    /// </summary>
    protected abstract void OnGenerationStarted() ;

    /// <summary>
    /// When overridden in a derived class, this method is called after each routing result is processed.
    /// </summary>
    /// <param name="routingTargets">Processed routing target, given by <see cref="RoutingTargetGroups"/> property.</param>
    /// <param name="result">Routing result.</param>
    protected abstract void OnRoutingTargetProcessed( IReadOnlyCollection<TAutoRoutingTarget> routingTargets, MergedAutoRoutingResult result ) ;

    /// <summary>
    /// When overridden in a derived class, this method is called after all route generations. It is good to postprocess for an execution.
    /// </summary>
    protected abstract void OnGenerationFinished() ;

    /// <summary>
    /// Execute generation routes.
    /// </summary>
    /// <param name="progressData">Progress data which is notified the status.</param>
    public void Execute( IProgressData? progressData )
    {
      progressData?.ThrowIfCanceled() ;

      using ( progressData?.Reserve( 0.05 ) ) {
        OnGenerationStarted() ;
      }

      progressData?.ThrowIfCanceled() ;

      var targetToIndex = new Dictionary<TAutoRoutingTarget, int>() ;
      var targetMergers = new List<AutoRoutingTargetMerger<TAutoRoutingTarget>>() ;
      RoutingTargetGroups.ForEach( ( group, index ) =>
      {
        group.ForEach( target => targetToIndex.Add( target, index ) ) ;
        targetMergers.Add( new AutoRoutingTargetMerger<TAutoRoutingTarget>( group.Count ) ) ;
      } ) ;

      using ( var mainProgress = progressData?.Reserve( 0.9 ) ) {
        var reporter = ( null == mainProgress ? new ProgressReporter() : new ProgressReporter( mainProgress ) ) ;
        var token = mainProgress?.CancellationToken ?? CancellationToken.None ;
        mainProgress.ForEach( targetToIndex.Count, ApiForAutoRouting.Execute( StructureGraph, RoutingTargetGroups.SelectMany( group => group ), CollisionCheckTree, reporter, token ), item =>
        {
          var (src, result) = item ;
          if ( src is not TAutoRoutingTarget srcTarget ) throw new InvalidOperationException() ;

          var merger = targetMergers[ targetToIndex[ srcTarget ] ] ;
          if ( false == merger.Register( srcTarget, result ) ) throw new InvalidOperationException() ;

          if ( merger.IsFullfilled ) {
            OnRoutingTargetProcessed( merger.GetAutoRoutingTargets(), merger.GetAutoRoutingResult() ) ;
          }

          progressData?.ThrowIfCanceled() ;
        } ) ;
      }

      using ( progressData?.Reserve( 1 - progressData.Position ) ) {
        OnGenerationFinished() ;
      }

      progressData?.ThrowIfCanceled() ;
    }
  }
}