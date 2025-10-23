using System.Collections.Generic ;
using System.Linq ;
using System.Windows ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public partial class ShowDialog : Window
  {
    public ShowDialog( string allCommandName )
    {
      InitializeComponent() ;
      string[] commandNames = allCommandName.Split( '.' ) ;
      messageBlock.Text = allCommandName ;
    }
  }
}