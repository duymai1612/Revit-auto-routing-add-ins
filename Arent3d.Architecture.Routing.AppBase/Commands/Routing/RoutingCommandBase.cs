using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Threading ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Revit.UI ;
using Arent3d.Revit.UI.Forms ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase.Commands.Routing
{
  public abstract class RoutingCommandBase<TUIResult> : ExternalCommandBase<TUIResult>
  {
    protected override string GetTransactionName() => GetTransactionNameKey().GetAppStringByKeyOrDefault( "Routing" ) ;

    protected override ExternalCommandTransactionType TransactionType => ExternalCommandTransactionType.TransactionGroup ;

    protected override IDisposable? BeforeCommand( ExternalCommandData commandData, ElementSet elements )
    {
      SetRoutingExecutor( CreateRoutingExecutor( commandData.Application.ActiveUIDocument.Document, commandData.View ) ) ;
      return base.BeforeCommand( commandData, elements ) ;
    }

    protected override void AfterCommand( IDisposable? commandSpanResource )
    {
      SetRoutingExecutor( null ) ;
      base.AfterCommand( commandSpanResource ) ;
    }

    private RoutingExecutor? _routingExecutor ;
    private void SetRoutingExecutor( RoutingExecutor? routingExecutor ) => _routingExecutor = routingExecutor ;
    protected RoutingExecutor GetRoutingExecutor() => _routingExecutor ?? throw new InvalidOperationException() ;

    protected sealed override ExecutionResult Execute( Document document, TransactionWrapper transaction, TUIResult result )
    {
      var executor = GetRoutingExecutor() ;

      var executionResult = GenerateRoutes( document, executor, result ) ;
      if ( Result.Cancelled == executionResult.Result ) return ExecutionResult.Cancelled ;
      if ( Result.Failed == executionResult.Result ) return ExecutionResult.Failed ;

      // Avoid Revit bugs about reducer insertion.
      FixReducers( document, executor, executionResult.Value ) ;

      // execute after route command
      AfterRouteGenerated( document, executionResult.Value ) ;

      return ExecutionResult.Succeeded ;
    }

    protected override void AfterTransaction( Document document, Result result, TUIResult uiResult )
    {
      var executor = GetRoutingExecutor() ;

      if ( Result.Succeeded == result ) {
        if ( executor.HasDeletedElements ) {
          CommandUtils.AlertDeletedElements() ;
        }

        if ( executor.HasBadConnectors ) {
          CommandUtils.AlertBadConnectors( executor.GetBadConnectorSet() ) ;
        }
      }
    }

    protected override string? OnException( Exception e, TUIResult? uiResult )
    {
      CommandUtils.DebugAlertException( e ) ;
      return null ;
    }

    protected abstract RoutingExecutor CreateRoutingExecutor( Document document, View view ) ;

    private OperationResult<IReadOnlyCollection<Route>> GenerateRoutes( Document document, RoutingExecutor executor, TUIResult state )
    {
      return document.Transaction( "TransactionName.Commands.Routing.Common.Routing".GetAppStringByKeyOrDefault( "Routing" ), transaction =>
      {
        SetupFailureHandlingOptions( transaction, executor ) ;

        try {
          using var progress = ProgressBar.ShowWithNewThread( new CancellationTokenSource() ) ;
          progress.Message = "Routing..." ;

          var segments = GetRouteSegments( document, state ) ;
          return executor.Run( segments, progress ) ;
        }
        catch ( OperationCanceledException ) {
          return OperationResult<IReadOnlyCollection<Route>>.Cancelled ;
        }
      } ) ;
    }

    private static void FixReducers( Document document, RoutingExecutor executor, IEnumerable<Route> routes )
    {
      document.Transaction( "TransactionName.Commands.Routing.Common.Routing".GetAppStringByKeyOrDefault( "Routing" ), transaction =>
      {
        SetupFailureHandlingOptions( transaction, executor ) ;

        var routeNames = routes.Select( route => route.Name ).ToHashSet() ;
        foreach ( var curve in document.GetAllElements<MEPCurve>().Where( curve => curve.GetRouteName() is { } routeName && routeNames.Contains( routeName ) ) ) {
          FixCurveReducers( curve ) ;
        }

        return Result.Succeeded ;
      } ) ;

      static void FixCurveReducers( MEPCurve curve )
      {
        // Avoid Revit bugs about reducer insertion.
        foreach ( var connector in curve.GetConnectors().OfEnd() ) {
          var anotherConnectors = connector.GetConnectedConnectors().OfEnd().EnumerateAll() ;
          if ( 1 != anotherConnectors.Count ) continue ;

          var anotherConnector = anotherConnectors.First() ;
          if ( connector.HasSameShapeAndParameters( anotherConnector ) ) continue ;

          if ( false == ShakeShape( connector, anotherConnector ) ) continue ;
          return ;  // done
        }
      }

      static bool ShakeShape( Connector connector, Connector anotherConnector )
      {
        switch ( connector.Shape ) {
          case ConnectorProfileType.Oval :
          case ConnectorProfileType.Round :
          {
            var orgRadius = connector.Radius ;
            connector.Radius = anotherConnector.Radius ;
            connector.Radius = orgRadius ;
            return true ;
          }
          case ConnectorProfileType.Rectangular :
          {
            var orgWidth = connector.Width ;
            var orgHeight = connector.Height ;
            connector.Width = anotherConnector.Width ;
            connector.Height = anotherConnector.Height ;
            connector.Width = orgWidth ;
            connector.Height = orgHeight ;
            return true ;
          }
          default : return false ;
        }
      }
    }


    private static void SetupFailureHandlingOptions( Transaction transaction, RoutingExecutor executor )
    {
      if ( executor.CreateFailuresPreprocessor() is not { } failuresPreprocessor ) return ;
      
      transaction.SetFailureHandlingOptions( ModifyFailureHandlingOptions( transaction.GetFailureHandlingOptions(), failuresPreprocessor ) ) ;
    }

    private static FailureHandlingOptions ModifyFailureHandlingOptions( FailureHandlingOptions handlingOptions, IFailuresPreprocessor failuresPreprocessor )
    {
      return handlingOptions.SetFailuresPreprocessor( failuresPreprocessor ) ;
    }

    protected abstract string GetTransactionNameKey() ;

    /// <summary>
    /// Generate route segments to be auto-routed from UI state.
    /// </summary>
    /// <returns>Routing from-to records.</returns>
    protected abstract IReadOnlyCollection<(string RouteName, RouteSegment Segment)> GetRouteSegments( Document document, TUIResult state ) ;

    protected virtual void AfterRouteGenerated( Document document, IReadOnlyCollection<Route> executeResultValue )
    {
    }
  }

  public abstract class RoutingCommandBaseWithoutOperation : RoutingCommandBase<object?>
  {
    protected sealed override OperationResult<object?> OperateUI( ExternalCommandData commandData, ElementSet elements ) => new(Result.Succeeded) ;
    protected sealed override OperationResult<object?> ReOperateUI( ExternalCommandData commandData, ElementSet elements, object? lastUiResult, ExecutionResultType lastExecutionResultType ) => new(Result.Succeeded) ;

    protected sealed override IDisposable? BeforeTransaction( Document document, object? uiResult ) => BeforeTransaction( document ) ;
    protected sealed override void AfterTransaction( Document document, Result result, object? uiResult ) => AfterTransaction( document, result ) ;

    protected virtual IDisposable? BeforeTransaction( Document document ) => null ;

    protected virtual void AfterTransaction( Document document, Result result )
    {
    }

    protected sealed override IReadOnlyCollection<(string RouteName, RouteSegment Segment)> GetRouteSegments( Document document, object? state )
    {
      return GetRouteSegments( document ) ;
    }

    protected abstract IReadOnlyCollection<(string RouteName, RouteSegment Segment)> GetRouteSegments( Document document ) ;
  }

  public abstract class RoutingCommandBaseWithParam<TCommandParameter> : RoutingCommandBaseWithoutOperation, IExternalCommandWithParam<TCommandParameter> where TCommandParameter : class
  {
    private TCommandParameter? Parameter { get ; set ; }

    protected sealed override IDisposable? BeforeCommand( ExternalCommandData commandData, ElementSet elements )
    {
      if ( false == this.PopCommandParameter( out var param ) ) return null ;

      Parameter = param ;
      return BeforeCommand( Parameter, commandData, elements ) ;
    }

    protected sealed override void AfterCommand( IDisposable? commandSpanResource )
    {
      var param = Parameter ;
      Parameter = null ;

      AfterCommand( param!, commandSpanResource ) ;
    }

    protected virtual IDisposable? BeforeCommand( TCommandParameter param, ExternalCommandData commandData, ElementSet elements ) => null ;
    protected virtual void AfterCommand( TCommandParameter param, IDisposable? commandSpanResource )
    {
    }

    protected sealed override IDisposable? BeforeTransaction( Document document ) => BeforeTransaction( Parameter!, document ) ;
    protected sealed override void AfterTransaction( Document document, Result result ) => AfterTransaction( Parameter!, document, result ) ;

    protected virtual IDisposable? BeforeTransaction( TCommandParameter param, Document document ) => null ;

    protected virtual void AfterTransaction( TCommandParameter param, Document document, Result result )
    {
    }

    protected sealed override IReadOnlyCollection<(string RouteName, RouteSegment Segment)> GetRouteSegments( Document document ) => GetRouteSegments( Parameter!, document ) ;

    protected abstract IReadOnlyCollection<(string RouteName, RouteSegment Segment)> GetRouteSegments( TCommandParameter param, Document document ) ;
  }
}