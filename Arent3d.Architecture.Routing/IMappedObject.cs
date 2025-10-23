using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// An object which implements <see cref="IMappedObject{TRevitElement}"/> is managed by a <see cref="ObjectMapper{TMapper,TRevitElement,TRoutingObject}"/>.
  /// </summary>
  /// <typeparam name="TRevitElement"></typeparam>
  public interface IMappedObject<out TRevitElement> where TRevitElement : Element
  {
    /// <summary>
    /// Returns a Revit object which this object is based.
    /// </summary>
    TRevitElement Element { get ; }
  }
}