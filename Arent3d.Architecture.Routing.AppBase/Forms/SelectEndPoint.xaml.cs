using System ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using System.Collections.ObjectModel ;
using System.ComponentModel ;
using System.Runtime.CompilerServices ;
using System.Windows ;
using System.Linq ;
using Arent3d.Architecture.Routing.EndPoints ;
using Arent3d.Revit ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  /// <summary>
  /// Interaction logic for SelectConnector.xaml
  /// </summary>
  public partial class SelectEndPoint : Window
  {
    private readonly IEndPoint? _firstEndPoint ;

    public ObservableCollection<EndPointInfoClass> EndPointList { get ; } = new() ;

    public SelectEndPoint( Document document, IEndPoint[] points, IEndPoint? firstEndPoint = null )
    {
      InitializeComponent() ;

      _firstEndPoint = firstEndPoint ;

      foreach ( IEndPoint conn in points ) {
        EndPointList.Add( new EndPointInfoClass( document, conn ) ) ;
      }

      this.Left = 0 ;
      this.Top += 10 ;
    }

    public class EndPointInfoClass : INotifyPropertyChanged
    {
      private bool _isSelected = false ;

      public bool IsSelected
      {
        get => _isSelected ;
        set
        {
          _isSelected = value ;
          NotifyPropertyChanged() ;
        }
      }

      public event PropertyChangedEventHandler? PropertyChanged ;

      private IEndPoint? Pointer { get ; }
      private readonly string _label ;

      public EndPointInfoClass( Document document, IEndPoint point )
      {
        Pointer = point ;
        if ( point is IRouteBranchEndPoint routeEndPoint ) {
          _label = $"{routeEndPoint.DisplayTypeName} - {routeEndPoint.RouteName}" ;
        }
        else {
          var position = point.RoutingStartPosition ;
          _label = $"{Pointer.DisplayTypeName} - {GetDisplayValue( document, position.X )}, {GetDisplayValue( document, position.Y )}, {GetDisplayValue( document, position.Z )}" ;
        }
      }

      private static string GetDisplayValue( Document document, double value )
      {
        if ( DisplayUnit.METRIC == document.DisplayUnitSystem ) {
          return value.RevitUnitsToMeters().ToString( "F2" ) ;
        }
        else {
          return value.RevitUnitsToFeet().ToString( "F2" ) ;
        }
      }

      private void NotifyPropertyChanged( [CallerMemberName] string propertyName = "" )
      {
        PropertyChanged?.Invoke( this, new PropertyChangedEventArgs( propertyName ) ) ;
      }

      public IEndPoint? GetEndPoint()
      {
        if ( false == IsSelected ) return null ;

        return Pointer ;
      }

      public override string ToString() => _label ;
    }

    private void Button_Click( object sender, RoutedEventArgs e )
    {
      this.DialogResult = true ;
      this.Close() ;
    }

    public IEndPoint GetSelectedEndPoint()
    {
      return EndPointList.Select( cic => cic.GetEndPoint() ).NonNull().FirstOrDefault()! ;
    }
  }
}