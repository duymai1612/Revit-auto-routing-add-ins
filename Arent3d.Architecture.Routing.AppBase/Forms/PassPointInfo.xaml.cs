using System.Windows ;
using System.Windows.Controls ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public partial class PassPointInfo : UserControl
  {
    public static readonly DependencyProperty DisplayUnitSystemProperty = DependencyProperty.Register( "DisplayUnitSystem", typeof( DisplayUnit ), typeof( PassPointInfo ), new PropertyMetadata( DisplayUnit.IMPERIAL ) ) ;
    public static readonly DependencyProperty XPointProperty = DependencyProperty.Register( "XPoint", typeof( double ), typeof( PassPointInfo ), new PropertyMetadata( 0.0 ) ) ;
    public static readonly DependencyProperty YPointProperty = DependencyProperty.Register( "YPoint", typeof( double ), typeof( PassPointInfo ), new PropertyMetadata( 0.0 ) ) ;
    public static readonly DependencyProperty ZPointProperty = DependencyProperty.Register( "ZPoint", typeof( double ), typeof( PassPointInfo ), new PropertyMetadata( 0.0 ) ) ;

    public DisplayUnit DisplayUnitSystem
    {
      get { return (DisplayUnit)GetValue( DisplayUnitSystemProperty ) ; }
      set { SetValue( DisplayUnitSystemProperty, value ) ; }
    }

    public double XPoint
    {
      get { return (double)GetValue( XPointProperty ) ; }
      set { SetValue( XPointProperty, value ) ; }
    }

    public double YPoint
    {
      get { return (double)GetValue( YPointProperty ) ; }
      set { SetValue( YPointProperty, value ) ; }
    }

    public double ZPoint
    {
      get { return (double)GetValue( ZPointProperty ) ; }
      set { SetValue( ZPointProperty, value ) ; }
    }

    public PassPointInfo()
    {
      InitializeComponent() ;
    }
  }
}