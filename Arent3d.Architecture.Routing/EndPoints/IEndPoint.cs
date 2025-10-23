using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.EndPoints
{
  public interface IEndPoint
  {
    /// <summary>
    /// Type name of this end point.
    /// </summary>
    string TypeName { get ; }

    /// <summary>
    /// Type name of this end point (displayed).
    /// </summary>
    string DisplayTypeName { get ; }

    /// <summary>
    /// A specifier which determine the equality of an end point.
    /// </summary>
    EndPointKey Key { get ; }

    /// <summary>
    /// Returns whether end point can be replaced into the other one.
    /// </summary>
    bool IsReplaceable { get ; }

    /// <summary>
    /// Returns whether this end point can have only one MEPCurve from this end point.
    /// </summary>
    bool IsOneSided { get ; }

    /// <summary>
    /// Stringified parameters to be saved or exported.
    /// </summary>
    string ParameterString { get ; }

    /// <summary>
    /// Returns a reference connector which is used to collect routing information.
    /// </summary>
    /// <returns></returns>
    Connector? GetReferenceConnector() ;

    /// <summary>
    /// Returns the start position of routing.
    /// </summary>
    XYZ RoutingStartPosition { get ; }

    /// <summary>
    /// Returns the flow vector.
    /// </summary>
    XYZ GetRoutingDirection( bool isFrom ) ;

    /// <summary>
    /// Returns whether this end point can start routing with specified direction.
    /// </summary>
    /// <param name="isFrom"></param>
    /// <returns></returns>
    bool HasValidElement( bool isFrom ) ;

    /// <summary>
    /// Gets the specified diameter when the end point has it.
    /// </summary>
    /// <returns>Diameter, if specified.</returns>
    double? GetDiameter() ;

    /// <summary>
    /// Returns level id which this end point is placed on. 
    /// </summary>
    /// <param name="document"></param>
    /// <returns></returns>
    ElementId GetLevelId( Document document ) ;

    /// <summary>
    /// Returns the required minimum straight length.
    /// </summary>
    double GetMinimumStraightLength( double edgeDiameter, bool isFrom ) ;

    /// <summary>
    /// Gets a parent route when the end point is dependent to it.
    /// </summary>
    /// <returns>Route which this end point is connecting.</returns>
    Route? ParentRoute() ;
    /// <summary>
    /// Gets a parent sub-route when the end point is dependent to it.
    /// </summary>
    /// <returns>SubRoute which this end point is connecting.</returns>
    SubRoute? ParentSubRoute() ;

    /// <summary>
    /// Generate an element if needed.
    /// </summary>
    /// <param name="routeName">Owner route name.</param>
    /// <returns>True if newly generated.</returns>
    bool GenerateInstance( string routeName ) ;

    /// <summary>
    /// Erase an element from document.
    /// </summary>
    /// <returns>True if erased.</returns>
    bool EraseInstance() ;

    void Accept( IEndPointVisitor visitor ) ;
    T Accept<T>( IEndPointVisitor<T> visitor ) ;
  }

  public interface IRealEndPoint : IEndPoint
  {
  }

  public interface IEndPointOfPassPoint : IEndPoint
  {
    ElementId PassPointId { get ; }
  }

  public interface IRouteBranchEndPoint : IEndPoint
  {
    string RouteName { get ; }
  }
}