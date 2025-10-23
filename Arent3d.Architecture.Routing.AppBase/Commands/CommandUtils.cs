using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase.Commands
{
  public static class CommandUtils
  {
    public static void AlertDeletedElements()
    {
      TaskDialog.Show( "Dialog.Commands.Routing.Common.Title.Error".GetAppStringByKeyOrDefault( null ), "Dialog.Commands.Routing.Common.Body.Error.DeletedSomeFailedElements".GetAppStringByKeyOrDefault( null ) ) ;
    }

    public static void AlertBadConnectors( IEnumerable<Connector[]> badConnectorSet )
    {
      var body = string.Format( "Dialog.Commands.Routing.Common.Body.Error.FittingCannotBeInserted".GetAppStringByKeyOrDefault( null ), "・" + string.Join( "\n・", badConnectorSet.Select( GetConnectorInfo ) ) ) ;
      TaskDialog.Show( "Dialog.Commands.Routing.Common.Title.Error".GetAppStringByKeyOrDefault( null ), body ) ;
    }

    public static void DebugAlertException( Exception e )
    {
#if DEBUG
      TaskDialog.Show( "Debug", e.ToString() ) ;
#else
      TaskDialog.Show( "Dialog.Commands.Routing.Common.Title.Error".GetAppStringByKeyOrDefault( null ), "Dialog.Commands.Routing.Common.Body.Error.ExceptionOccured".GetAppStringByKeyOrDefault( null ) ) ;
#endif
    }

    private static string GetConnectorInfo( Connector[] connectorSet )
    {
      var connectionType = connectorSet.Length switch { 2 => "Elbow", 3 => "Tee", 4 => "Cross", _ => throw new ArgumentException() } ;
      var connector = connectorSet.FirstOrDefault( c => c.IsValidObject ) ;
      var coords = ( null != connector ) ? GetCoordValue( connector.Owner.Document, connector.Origin ) : "(Deleted connectors)" ;
      return $"[{connectionType}] {coords}" ;
    }

    private static string GetCoordValue( Document document, XYZ pos )
    {
      return document.DisplayUnitSystem switch
      {
        DisplayUnit.METRIC => $"({pos.X.RevitUnitsToMeters()}, {pos.Y.RevitUnitsToMeters()}, {pos.Z.RevitUnitsToMeters()})",
        _ => $"({pos.X.RevitUnitsToFeet()}, {pos.Y.RevitUnitsToFeet()}, {pos.Z.RevitUnitsToFeet()})",
      } ;
    }
  }
}