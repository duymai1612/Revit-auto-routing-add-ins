using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Windows ;
using Autodesk.Revit.DB ;
using System.Collections.ObjectModel ;
using System.ComponentModel ;
using System.Runtime.CompilerServices ;
using Arent3d.Revit ;
using Arent3d.Utility ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  /// <summary>
  /// getLevel.xaml の相互作用ロジック
  /// </summary>
  public partial class GetLevel : Window
  {
    public ObservableCollection<LevelInfo> LevelList { get ; }

    public GetLevel( Document doc )
    {
      InitializeComponent() ;

      LevelList = new ObservableCollection<LevelInfo>( doc.GetAllElements<Level>().OfCategory( BuiltInCategory.OST_Levels ).Select( ToLevelInfo ).OrderBy( l => l.Elevation ) ) ;
    }

    private static LevelInfo ToLevelInfo( Level level )
    {
      return new LevelInfo { Elevation = level.Elevation, LevelId = level.Id, IsSelected = false, LevelName = level.Name } ;
    }

    public IReadOnlyCollection<(ElementId Id, string Name)> GetSelectedLevels()
    {
      if ( true != DialogResult ) return Array.Empty<(ElementId, string)>() ;

      return LevelList.Where( level => level.IsSelected ).Select( level => ( Id: level.LevelId, Name: level.LevelName ) ).EnumerateAll() ;
    }

    public class LevelInfo : INotifyPropertyChanged
    {
      public string LevelName { get ; init ; } = string.Empty ;
      public ElementId LevelId { get ; init ; } = ElementId.InvalidElementId ;
      public double Elevation { get ; init ; }
      private bool _isSelected ;

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

      private void NotifyPropertyChanged( [CallerMemberName] string propertyName = "" )
      {
        PropertyChanged?.Invoke( this, new PropertyChangedEventArgs( propertyName ) ) ;
      }
    }

    private void CheckAll( object sender, RoutedEventArgs e )
    {
      SelectAll( true ) ;
    }

    private void UncheckAll( object sender, RoutedEventArgs e )
    {
      SelectAll( false ) ;
    }

    private void ToggleAll( object sender, RoutedEventArgs e )
    {
      SelectToggle() ;
    }

    private void SelectButton_Click( object sender, RoutedEventArgs e )
    {
      this.DialogResult = true ;
      this.Close() ;
    }

    private void SelectAll( bool select )
    {
      LevelList.ForEach( level => level.IsSelected = select ) ;
    }

    private void SelectToggle()
    {
      LevelList.ForEach( level => level.IsSelected = ! level.IsSelected ) ;
    }
  }
}