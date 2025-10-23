using System ;
using System.Collections.Generic ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Revit.UI ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Structure ;
using Autodesk.Revit.UI ;
using Autodesk.Revit.UI.Selection ;

namespace Arent3d.Architecture.Routing.AppBase
{
  public sealed class EndPointPicker : MustBeDisposed
  {
    private readonly UIDocument _uiDocument ;
    private readonly TempColor _routeTempColor ;
    private readonly TempColor _endPointsTempColor ;
    private readonly Dictionary<ElementId, IEndPoint> _endPointsByElementId = new() ;

    public EndPointPicker( UIDocument uiDocument, IEnumerable<ElementId> routeElements, IEnumerable<IEndPoint> endPoints )
    {
      _uiDocument = uiDocument ;

      var document = uiDocument.Document ;
      var symbol = document.GetFamilySymbols( RoutingFamilyType.ConnectorPoint ).FirstOrDefault() ?? throw new InvalidOperationException() ;
      _routeTempColor = new TempColor( uiDocument.ActiveView, new Color( 0, 0, 255 ) ) ;
      _endPointsTempColor = new TempColor( uiDocument.ActiveView, new Color( 255, 0, 255 ) ) ;
      document.Transaction( "TransactionName.Commands.Routing.Common.PickEndPointOverSubRoute".GetAppStringByKeyOrDefault( null ), t =>
      {
        if ( false == symbol.IsActive ) symbol.Activate() ;

        _routeTempColor.AddRange( routeElements ) ;

        foreach ( var endPoint in endPoints ) {
          var instance = document.Create.NewFamilyInstance( endPoint.RoutingStartPosition, symbol, StructuralType.NonStructural ) ;
          _endPointsTempColor.Add( instance.Id );
          _endPointsByElementId.Add( instance.Id, endPoint ) ;
        }

        return Result.Succeeded ;
      } ) ;
    }

    protected override void Finally()
    {
      var document = _uiDocument.Document ;

      document.Transaction( "TransactionName.Commands.Routing.Common.EndPickEndPointOverSubRoute".GetAppStringByKeyOrDefault( null ), t =>
      {
        _endPointsTempColor.Dispose() ;
        _routeTempColor.Dispose() ;
        document.Delete( _endPointsByElementId.Keys ) ;

        return Result.Succeeded ;
      } ) ;
    }

    public IEndPoint? Pick()
    {
      var elm = _uiDocument.Selection.PickObject( ObjectType.Element, new GetElementFilter( _endPointsByElementId ), "Dialog.Commands.Routing.PickRouting.PickEndPointOverSubRoute".GetAppStringByKeyOrDefault( null ) ) ;
      if ( null == elm ) return null ;
      if ( false == _endPointsByElementId.TryGetValue( elm.ElementId, out var endPoint ) ) return null ;

      return endPoint ;
    }

    private class GetElementFilter : ISelectionFilter
    {
      private readonly IReadOnlyDictionary<ElementId, IEndPoint> _dic ;

      public GetElementFilter( IReadOnlyDictionary<ElementId, IEndPoint> dic ) => _dic = dic ;
      public bool AllowElement( Element elem ) => _dic.ContainsKey( elem.Id ) ;
      public bool AllowReference( Reference reference, XYZ position ) => _dic.ContainsKey( reference.ElementId ) ;
    }
  }
}