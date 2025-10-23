using System ;

namespace Arent3d.Architecture.Routing.Utils
{
  public static class StringUtils
  {

    public static string DefaultIfBlank( string? text, string defaultValue )
    {
      return string.IsNullOrWhiteSpace( text ) ? defaultValue : text! ;
    }
    
  }
}