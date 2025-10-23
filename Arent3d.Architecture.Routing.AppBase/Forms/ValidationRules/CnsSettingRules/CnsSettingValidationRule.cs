using System.Globalization ;
using System.Windows.Controls ;

namespace Arent3d.Architecture.Routing.AppBase.Forms.ValidationRules.CnsSettingRules
{
  public class CnsSettingValidationRule : ValidationRule
  {
    public override ValidationResult Validate( object value, CultureInfo cultureInfo )
    {
      if ( value == null || string.IsNullOrWhiteSpace( value.ToString() ) ) {
        return new ValidationResult( false, "Value is required." ) ;
      }

      return ValidationResult.ValidResult ;
    }
  }
}
