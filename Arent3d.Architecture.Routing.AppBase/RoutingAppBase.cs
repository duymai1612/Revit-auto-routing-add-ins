using System ;
using System.Collections.Generic ;
using System.Reflection ;
using Arent3d.Architecture.Routing.AppBase.Updater ;
using Arent3d.Revit ;
using Arent3d.Revit.UI ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Events ;
using Autodesk.Revit.UI ;
using Autodesk.Revit.UI.Events ;

namespace Arent3d.Architecture.Routing.AppBase
{
  /// <summary>
  /// Entry point of auto routing application. This class calls UI initializers.
  /// </summary>
  public abstract class RoutingAppBase : ExternalApplicationBase
  {
    protected override string GetLanguageDirectoryPath()
    {
      return GetDefaultLanguageDirectoryPath( Assembly.GetExecutingAssembly() ) ;
    }

    protected override IAppUIBase? CreateAppUI( UIControlledApplication application )
    {
      // Validate StorableBase classes
      StorableBase.ValidateAllStorableClassDefinitions( typeof( AppInfo ).Assembly ) ;

      return null ;
    }

    protected override void RegisterEvents( UIControlledApplication application )
    {
    }

    protected override void UnregisterEvents( UIControlledApplication application )
    {
    }

    protected override void OnDocumentListenStarted( Document document )
    {
      DocumentMapper.Register( document ) ;
    }

    protected override void OnDocumentListenFinished( Document document )
    {
      DocumentMapper.Unregister( document ) ;
    }

    protected override void OnDocumentChanged( Document document, DocumentChangedEventArgs e )
    {
    }

    protected override void OnApplicationViewChanged( Document document, ViewActivatedEventArgs e )
    {
    }

    protected override IEnumerable<IDocumentUpdateListener> GetUpdateListeners() => Array.Empty<IDocumentUpdateListener>() ;
  }
}