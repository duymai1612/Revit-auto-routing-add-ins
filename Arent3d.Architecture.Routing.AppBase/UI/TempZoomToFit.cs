using System ;
using Arent3d.Revit.UI ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.UI ;

namespace Arent3d.Architecture.Routing.AppBase.UI
{
  public class TempZoomToFit : MustBeDisposed
  {
    private readonly UIView? _uiView ;
    private readonly View? _view ;
    private readonly XYZ? _viewCenter ;
    private readonly XYZ? _upDirection ;
    private readonly XYZ? _forwardDirection ;
    private readonly XYZ? _corner1 ;
    private readonly XYZ? _corner2 ;
    private readonly double _scale ;
    
    public TempZoomToFit( UIDocument uiDocument )
    {
      _view = uiDocument.ActiveView ;
      _uiView = uiDocument.GetActiveUIView() ;

      if ( null != _uiView ) {
        if ( _view is View3D view3D ) {
          ( _viewCenter, _upDirection, _forwardDirection, _scale ) = GetCurrentCamera( _uiView, view3D ) ;
        }
        else if ( null != _view ) {
          var corners = _uiView.GetZoomCorners() ;
          _corner1 = corners[ 0 ] ;
          _corner2 = corners[ 1 ] ;
        }
      }
    }

    public void ZoomToFit()
    {
      _uiView?.ZoomToFit() ;
    }

    private static (XYZ ViewCenter, XYZ UpDirection, XYZ ForwardDirection, double Scale) GetCurrentCamera( UIView uiView, View3D view3D )
    {
      var corners = uiView.GetZoomCorners() ;
      var corner1 = corners[ 0 ] ;
      var corner2 = corners[ 1 ] ;

      var viewCenter = 0.5 * ( corner1 + corner2 ) ;
      var diagonalVector = corner1 - corner2 ;

      var viewOrientation3D = view3D.GetOrientation() ;
      var upDirection = viewOrientation3D.UpDirection ;
      var forwardDirection = viewOrientation3D.ForwardDirection ;
      var rightDirection = forwardDirection.CrossProduct( upDirection ) ;

      var height = Math.Abs( diagonalVector.DotProduct( upDirection ) ) ;
      var width = Math.Abs( diagonalVector.DotProduct( rightDirection ) ) ;

      var scale = 0.5 * Math.Min( height, width ) ;

      return ( viewCenter, upDirection, forwardDirection, scale ) ;
    }

    protected override void Finally()
    {
      if ( null == _uiView ) return ;

      if ( _view is View3D view3D ) {
        var orientation = new ViewOrientation3D( _viewCenter, _upDirection, _forwardDirection ) ;
        view3D.SetOrientation( orientation ) ;

        var corner1 = _viewCenter + _upDirection * _scale - view3D.RightDirection * _scale ;
        var corner2 = _viewCenter - _upDirection * _scale + view3D.RightDirection * _scale ;

        _uiView.ZoomAndCenterRectangle( corner1, corner2 ) ;
      }
      else if ( null != _corner1 && null != _corner2 ) {
        _uiView.ZoomAndCenterRectangle( _corner1, _corner2 ) ;
      }
    }
  }
}