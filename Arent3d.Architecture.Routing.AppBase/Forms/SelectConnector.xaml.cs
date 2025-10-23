using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using System ;
using System.Collections.Generic ;
using System.Collections.ObjectModel ;
using System.ComponentModel ;
using System.Runtime.CompilerServices ;
using System.Windows ;
using System.Linq ;
using Arent3d.Revit ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  /// <summary>
  /// Interaction logic for SelectConnector.xaml
  /// </summary>
  public partial class SelectConnector : Window
  {
    private readonly Connector? _firstConnector ;

    public ObservableCollection<ConnectorInfoClass> ConnectorList { get ; } = new() ;

    public SelectConnector( Element element, Connector? firstConnector, AddInType addInType )
    {
      InitializeComponent() ;

      _firstConnector = firstConnector ;

      if ( element is FamilyInstance familyInstance ) {
        foreach ( var (conn, connElm) in CreateConnector.GetConnectorAndConnectorElementPair( familyInstance, addInType ) ) {
          ConnectorList.Add( new ConnectorInfoClass( familyInstance, connElm, conn, _firstConnector ) ) ;
        }
      }
      else if ( element is MEPCurve curve ) {
        foreach ( var c in curve.GetConnectors().Where( c => ConnectorPicker.IsTargetConnector( c, addInType ) ) ) {
          ConnectorList.Add( new ConnectorInfoClass( c, _firstConnector ) ) ;
        }
      }

      ConnectorList.Add( new ConnectorInfoClass( element ) ) ;

      this.Left = 0 ;
      this.Top += 10 ;
    }

    public class ConnectorInfoClass : INotifyPropertyChanged
    {
      public bool IsEnabled { get ; }

      private bool _isSelected = false ;

      public bool IsSelected
      {
        get => _isSelected ;
        set
        {
          if ( false == IsEnabled ) return ;

          _isSelected = value ;
          NotifyPropertyChanged() ;
        }
      }

      public event PropertyChangedEventHandler? PropertyChanged ;

      private Element Element { get ; }

      private Connector? Connector { get ; }
      private ConnectorElement? ConnectorElement { get ; }

      /// <summary>
      /// ConnectorInfo for the center of an element.
      /// </summary>
      /// <param name="element">Instance.</param>
      public ConnectorInfoClass( Element element )
      {
        Element = element ;
        Connector = null ;
        ConnectorElement = null ;

        IsEnabled = true ;
      }

      public ConnectorInfoClass( FamilyInstance familyInstance, ConnectorElement connectorElement, Connector connector, Connector? firstConnector )
      {
        Element = familyInstance ;
        Connector = connector ;
        ConnectorElement = connectorElement ;

        IsEnabled = CreateConnector.IsEnabledConnector( connector, firstConnector ) ;
      }

      public ConnectorInfoClass( Connector connector, Connector? firstConnector )
      {
        Element = connector.Owner ;
        Connector = connector ;
        ConnectorElement = null ;

        IsEnabled = CreateConnector.IsEnabledConnector( connector, firstConnector ) ;
      }

      private void NotifyPropertyChanged( [CallerMemberName] string propertyName = "" )
      {
        PropertyChanged?.Invoke( this, new PropertyChangedEventArgs( propertyName ) ) ;
      }

      public override string ToString()
      {
        if ( null != ConnectorElement ) {
          return $"{ConnectorElement.Name} - φ {ConnectorElement.Radius.RevitUnitsToMillimeters() * 2} mm - {ConnectorElement.get_Parameter( BuiltInParameter.RBS_PIPE_FLOW_DIRECTION_PARAM )?.AsValueString()}" ;
        }
        else if ( null != Connector ) {
          return $"{Connector.Id} - φ {Connector.Radius.RevitUnitsToMillimeters() * 2} mm - {Connector.Direction}" ;
        }
        else {
          return "Origin of this element" ;
        }
      }

      public Connector? GetConnector()
      {
        if ( false == IsEnabled || false == IsSelected ) return null ;

        if ( null != Connector ) return Connector ;

        return null ;
      }
    }

    private void Button_Click( object sender, RoutedEventArgs e )
    {
      this.DialogResult = true ;
      this.Close() ;
    }

    public Connector? GetSelectedConnector()
    {
      return ConnectorList.Select( cic => cic.GetConnector() ).NonNull().FirstOrDefault() ;
    }
  }
}