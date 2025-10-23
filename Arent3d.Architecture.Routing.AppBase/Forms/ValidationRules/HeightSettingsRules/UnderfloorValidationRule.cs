using System.Globalization ;
using System.Windows.Controls ;

namespace Arent3d.Architecture.Routing.AppBase.Forms.ValidationRules.HeightSettingsRules
{
  internal class UnderfloorValidationRule : ValidationRule
  {
    private const int MAX_VALUE = 999999 ;

    public override ValidationResult Validate( object value, CultureInfo cultureInfo )
    {
      if ( value == null || string.IsNullOrWhiteSpace( value.ToString() ) ) {
        return new ValidationResult( false, "Value is required." ) ;
      }

      int proposedValue ;
      if ( ! int.TryParse( value.ToString(), out proposedValue ) ) {
        return new ValidationResult( false, "Value must be an integer value." ) ;
      }

      if ( proposedValue > MAX_VALUE ) {
        return new ValidationResult( false, $"Value must be less than or equal to {MAX_VALUE}." ) ;
      }


      return ValidationResult.ValidResult ;
    }
  }
}