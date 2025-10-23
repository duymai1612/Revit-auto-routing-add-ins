using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.CollisionTree ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Architecture.Routing.FittingSizeCalculators ;
using Arent3d.Architecture.Routing.StorableCaches ;
using Arent3d.Revit ;
using Arent3d.Revit.UI ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.Exceptions ;
using MathLib ;
using InvalidOperationException = System.InvalidOperationException ;

namespace Arent3d.Architecture.Routing.AppBase
{
  /// <summary>
  /// Routing execution object.
  /// </summary>
  public abstract class RoutingExecutor
  {
    private readonly PipeSpecDictionary _pipeSpecDictionary ;
    protected Document Document { get ; }
    public IFittingSizeCalculator FittingSizeCalculator { get ; }
    private readonly List<Connector[]> _badConnectors = new() ;

    /// <summary>
    /// Generates a routing execution object.
    /// </summary>
    /// <param name="document"></param>
    /// <param name="view"></param>
    /// <param name="fittingSizeCalculator"></param>
    protected RoutingExecutor( Document document, View view, IFittingSizeCalculator fittingSizeCalculator )
    {
      Document = document ;
      FittingSizeCalculator = fittingSizeCalculator ;
      _pipeSpecDictionary = new PipeSpecDictionary( document, fittingSizeCalculator ) ;
      CollectRacks( document, view, GetRackFamilyInstances() ) ;
    }

    protected abstract IEnumerable<FamilyInstance> GetRackFamilyInstances() ;

    private static void CollectRacks( Document document, View view, IEnumerable<FamilyInstance> rackFamilyInstances )
    {
      const double beamInterval = 6.0 ; // TODO
      const double sideBeamWidth = 0.2 ; // TODO
      const double sideBeamHeight = 0.2 ; // TODO
      var racks = DocumentMapper.Get( document ).RackCollection ;

      racks.Clear() ;
      foreach ( var familyInstance in rackFamilyInstances ) {
        var (min, max) = familyInstance.get_BoundingBox( view ).To3dRaw() ;

        racks.AddRack( new Rack.Rack( new Box3d( min, max ), beamInterval, sideBeamWidth, sideBeamHeight ) { IsMainRack = true } ) ;
      }

      racks.CreateLinkages() ;
    }

    /// <summary>
    /// Whether some connectors between ducts which elbows, tees or crosses could not be inserted. 
    /// </summary>
    public bool HasBadConnectors => ( 0 < _badConnectors.Count ) ;

    /// <summary>
    /// Returns connectors between ducts which elbows, tees or crosses could not be inserted. 
    /// </summary>
    /// <returns></returns>
    public IReadOnlyCollection<Connector[]> GetBadConnectorSet() => _badConnectors ;

    /// <summary>
    /// Execute routing for the passed routing records.
    /// </summary>
    /// <param name="fromToList">Routing from-to records.</param>
    /// <param name="progressData">Progress bar.</param>
    /// <returns>Result of execution.</returns>
    public OperationResult<IReadOnlyCollection<Route>> Run( IReadOnlyCollection<(string RouteName, RouteSegment Segment)> fromToList, IProgressData? progressData = null )
    {
      try {
        IReadOnlyCollection<Route> routes ;
        using ( var p = progressData?.Reserve( 0.01 ) ) {
          routes = ConvertToRoutes( fromToList, p ) ;
        }

        var domainRoutes = GroupByDomain( routes ) ;
        foreach ( var (domain, routesOfDomain) in domainRoutes ) {
          using var p = progressData?.Reserve( 0.99 * routesOfDomain.Count / routes.Count ) ;
          ExecuteRouting( domain, routesOfDomain, p ) ;
        }

        return new OperationResult<IReadOnlyCollection<Route>>( routes ) ;
      }
      catch ( OperationCanceledException ) {
        return OperationResult<IReadOnlyCollection<Route>>.Cancelled ;
      }
    }

    private static IEnumerable<(Domain, IReadOnlyCollection<Route>)> GroupByDomain( IReadOnlyCollection<Route> routes )
    {
      var dic = new Dictionary<Domain, List<Route>>() ;

      foreach ( var route in routes ) {
        var domain = route.Domain ;
        if ( false == IsRoutingDomain( domain ) ) continue ;

        if ( false == dic.TryGetValue( domain, out var list ) ) {
          list = new List<Route>() ;
          dic.Add( domain, list ) ;
        }

        list.Add( route ) ;
      }

      return dic.Select( pair => ( pair.Key, (IReadOnlyCollection<Route>) pair.Value ) ) ;
    }

    private static bool IsRoutingDomain( Domain domain )
    {
      return domain switch
      {
        Domain.DomainHvac => true,
        Domain.DomainPiping => true,
        Domain.DomainCableTrayConduit => true,
        Domain.DomainElectrical => true,
        _ => false,
      } ;
    }

    public MEPSystemPipeSpec GetMEPSystemPipeSpec( SubRoute subRoute ) => _pipeSpecDictionary.GetMEPSystemPipeSpec( subRoute ) ;

    private void ExecuteRouting( Domain domain, IReadOnlyCollection<Route> routes, IProgressData? progressData )
    {
      progressData?.ThrowIfCanceled() ;
      
      ICollisionCheckTargetCollector collector ;
      using ( progressData?.Reserve( 0.05 ) ) {
        collector = CreateCollisionCheckTargetCollector( domain, routes ) ;
      }

      progressData?.ThrowIfCanceled() ;
      
      RouteGenerator generator ;
      using ( progressData?.Reserve( 0.02 ) ) {
        generator = CreateRouteGenerator( routes, Document, collector ) ;
      }

      progressData?.ThrowIfCanceled() ;
      
      using ( var generatorProgressData = progressData?.Reserve( 1 - progressData.Position ) ) {
        generator.Execute( generatorProgressData ) ;
      }

      RegisterBadConnectors( generator.GetBadConnectorSet() ) ;

      routes.ForEach( r => r.Save() ) ;
    }

    protected abstract RouteGenerator CreateRouteGenerator( IReadOnlyCollection<Route> routes, Document document, ICollisionCheckTargetCollector collector ) ;

    protected abstract ICollisionCheckTargetCollector CreateCollisionCheckTargetCollector( Domain domain, IReadOnlyCollection<Route> routesInType ) ;

    /// <summary>
    /// Converts routing from-to records to routing objects.
    /// </summary>
    /// <param name="fromToList">Routing from-to records.</param>
    /// <param name="progressData">Progress bar.</param>
    /// <returns>Routing objects</returns>
    private IReadOnlyCollection<Route> ConvertToRoutes( IReadOnlyCollection<(string RouteName, RouteSegment Segment)> fromToList, IProgressData? progressData )
    {
      var oldRoutes = RouteCache.Get( Document ) ;

      var dic = new Dictionary<string, (Route, List<RouteSegment>)>() ;
      var result = new List<Route>() ;

      var parents = new HashSet<Route>() ;
      progressData.ForEach( fromToList, tuple =>
      {
        var (routeName, segment) = tuple ;

        if ( false == dic.TryGetValue( routeName, out var routeAndNewSegments ) ) {
          var route = oldRoutes.FindOrCreate( routeName ) ;

          routeAndNewSegments = ( route, new List<RouteSegment>() ) ;
          dic.Add( routeName, routeAndNewSegments ) ;
          result.Add( route ) ;
        }

        if ( segment.FromEndPoint.ParentRoute() is { } fromParent ) {
          parents.UnionWith( fromParent.GetAllRelatedBranches() ) ;
        }

        if ( segment.ToEndPoint.ParentRoute() is { } toParent ) {
          parents.UnionWith( toParent.GetAllRelatedBranches() ) ;
        }

        routeAndNewSegments.Item2.Add( segment ) ;

        progressData?.ThrowIfCanceled() ;
      } ) ;

      foreach ( var (route, segments) in dic.Values ) {
        route.Clear() ;
        segments.ForEach( segment => route.RegisterSegment( segment ) ) ;
      }

      result.AddRange( parents.Where( p => false == dic.ContainsKey( p.RouteName ) ) ) ;

      return result ;
    }

    private void RegisterBadConnectors( IEnumerable<Connector[]> badConnectorSet )
    {
      _badConnectors.AddRange( badConnectorSet ) ;
    }

    public bool HasDeletedElements { get ; set ; }

    public abstract IFailuresPreprocessor CreateFailuresPreprocessor() ;
  }
}