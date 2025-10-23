using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public abstract class RoutingFailuresPreprocessorBase : IFailuresPreprocessor
  {
    private readonly RoutingExecutor _executor ;

    protected RoutingFailuresPreprocessorBase( RoutingExecutor executor )
    {
      _executor = executor ;
    }

    public FailureProcessingResult PreprocessFailures( FailuresAccessor failuresAccessor )
    {
      var document = failuresAccessor.GetDocument() ;

      var elementsToDelete = new HashSet<ElementId>() ;
      foreach ( var failure in failuresAccessor.GetFailureMessages() ) {
        var handler = new FailureMessageHandler( failuresAccessor, failure ) ;
        if ( PreprocessFailureMessage( handler ) ) continue ;

        foreach ( var elmId in failure.GetFailingElementIds() ) {
          if ( document.GetElementById<MEPCurve>( elmId ) is null ) continue ;
          elementsToDelete.Add( elmId ) ;
        }
      }

      if ( 0 < elementsToDelete.Count ) {
        _executor.HasDeletedElements = true ;
        failuresAccessor.DeleteElements( elementsToDelete.ToList() ) ;

        return FailureProcessingResult.ProceedWithCommit ;
      }
      else {
        return FailureProcessingResult.Continue ;
      }
    }

    protected abstract bool PreprocessFailureMessage( FailureMessageHandler handler ) ;

    protected class FailureMessageHandler
    {
      private FailuresAccessor FailuresAccessor { get ; }
      private FailureMessageAccessor FailureMessageAccessor { get ; }

      public FailureMessageHandler( FailuresAccessor failuresAccessor, FailureMessageAccessor failureMessageAccessor )
      {
        FailuresAccessor = failuresAccessor ;
        FailureMessageAccessor = failureMessageAccessor ;
      }

      public FailureDefinitionId FailureDefinitionId => FailureMessageAccessor.GetFailureDefinitionId() ;
      public IEnumerable<ElementId> GetFailingElementIds() => FailureMessageAccessor.GetFailingElementIds() ;

      public void DeleteWarning() => FailuresAccessor.DeleteWarning( FailureMessageAccessor ) ;
    }
  }
}