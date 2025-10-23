using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Updater
{
  public class RoutingUpdateListener : IDocumentUpdateListener
  {
    public Guid Guid { get ; } = new Guid( "{3BC9CB75-EB32-4E64-8879-2644CCB07A68}" ) ;

    private readonly IUIApplicationHolder _holder ;

    public string Name { get ; }
    public string Description { get ; }
    public ChangePriority ChangePriority { get ; }
    public DocumentUpdateListenType ListenType { get ; }

    public ElementFilter GetElementFilter()
    {
      //Please change this method to filter the target families
      ElementFilter filter = new ElementCategoryFilter( BuiltInCategory.OST_MechanicalEquipment ) ;
      return filter ;
    }

    public IEnumerable<ParameterProxy> GetListeningParameters() => throw new NotSupportedException() ;

    public RoutingUpdateListener( IUIApplicationHolder holder )
    {
      _holder = holder ;

      Name = this.ToString() ;
      Description = "Update location " ;
      ListenType = DocumentUpdateListenType.Deletion | DocumentUpdateListenType.Geometry ;
    }

    public void Execute( UpdaterData data )
    {
      if ( _holder.UiApp is not { } uiApp ) return ;
      var document = uiApp.ActiveUIDocument.Document ;

      var elementIds = uiApp.ActiveUIDocument.Selection.GetElementIds() ;
      if ( 1 != elementIds.Count ) return ;

      if ( document.GetElement( elementIds.First() ) is not { } movedElement ) return ;

      var filteredElementCollector = new FilteredElementCollector( document ) ;
      var elementList = filteredElementCollector.OfClass( typeof( FamilyInstance ) ).OfCategory( BuiltInCategory.OST_MechanicalEquipment ).ToList() ;
      foreach ( Element element in elementList ) {
        foreach ( Parameter parameter in element.Parameters ) {
          if ( parameter.Definition.Name == "LinkedInstanceId" ) {
            if ( parameter.AsString() == movedElement.ToString() ) {
              LocationPoint? LpF = element.Location as LocationPoint ;
              XYZ? epFrom = LpF?.Point ;

              LocationPoint? LpT = movedElement.Location as LocationPoint ;
              XYZ? epTo = LpT?.Point ;
              element.Location.Move( epTo - epFrom ) ;
            }
          }
        }
      }
    }
  }
}