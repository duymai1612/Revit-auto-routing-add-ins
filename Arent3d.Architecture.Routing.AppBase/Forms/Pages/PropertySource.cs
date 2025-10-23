using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Electrical ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  /// <summary>
  /// Property source for selected item's UI
  /// </summary>
  public abstract class PropertySource
  {
    public Document Document { get ; }

    protected PropertySource( Document doc )
    {
      Document = doc ;
    }
  }

  /// <summary>
  /// PropertySource for RouteItem, SubRouteItem
  /// </summary>
  public class RoutePropertySource : PropertySource
  {
    //Route
    public Route TargetRoute { get ; }
    public IReadOnlyCollection<SubRoute> TargetSubRoutes { get ; }

    public RouteProperties Properties { get ; }

    public RoutePropertySource( Document doc, IReadOnlyCollection<SubRoute> subRoutes ) : base( doc )
    {
      if ( 0 == subRoutes.Count ) throw new ArgumentException() ;

      TargetSubRoutes = subRoutes ;
      TargetRoute = subRoutes.First().Route ;

      Properties = new RouteProperties( subRoutes ) ;
    }
  }

  /// <summary>
  /// PropertySource for Connector
  /// </summary>
  public class ConnectorPropertySource : PropertySource
  {
    public Connector TargetConnector { get ; }
    public XYZ ConnectorTransform { get ; }

    public ConnectorPropertySource( Document doc, Connector connector ) : base( doc )
    {
      TargetConnector = connector ;
      ConnectorTransform = connector.Origin ;
    }
  }

  /// <summary>
  /// PropertySource for PassPoint
  /// </summary>
  public class PassPointPropertySource : PropertySource
  {
    public PassPointEndPoint PassPointEndPoint { get ; }
    public XYZ PassPointPosition { get ; }
    public XYZ PassPointDirection { get ; }

    public PassPointPropertySource( Document doc, PassPointEndPoint passPointEndPoint ) : base( doc )
    {
      PassPointEndPoint = passPointEndPoint ;

      PassPointPosition = passPointEndPoint.RoutingStartPosition ;
      PassPointDirection = passPointEndPoint.Direction ;
    }
  }

  /// <summary>
  /// PropertySource for TerminatePoint
  /// </summary>
  public class TerminatePointPropertySource : PropertySource
  {
    public TerminatePointEndPoint TerminatePointEndPoint { get ; }
    public XYZ TerminatePointPosition { get ; }
    public XYZ TerminatePointDirection { get ; }

    public ElementId LinkedElementId { get ; }
    public string? LinkedElementName { get ; }

    public TerminatePointPropertySource( Document doc, TerminatePointEndPoint terminatePointEndPoint ) : base( doc )
    {
      TerminatePointEndPoint = terminatePointEndPoint ;

      TerminatePointPosition = terminatePointEndPoint.RoutingStartPosition ;
      TerminatePointDirection = terminatePointEndPoint.Direction ;

      LinkedElementId = terminatePointEndPoint.LinkedInstanceId ;
      LinkedElementName = doc.GetElement( LinkedElementId )?.Name ;
    }
  }
}