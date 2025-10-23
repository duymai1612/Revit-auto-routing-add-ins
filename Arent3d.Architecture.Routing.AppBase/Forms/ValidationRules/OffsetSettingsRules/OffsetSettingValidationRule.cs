using System.Globalization ;
using System.Windows.Controls ;

namespace Arent3d.Architecture.Routing.AppBase.Forms.ValidationRules.OffsetSettingsRules
{
  internal class OffsetSettingValidationRule : ValidationRule
  {
    private const double MAX_VALUE = 999999 ;

    public override ValidationResult Validate( object value, CultureInfo cultureInfo )
    {
      if ( string.IsNullOrWhiteSpace( value.ToString() ) ) {
        return new ValidationResult( false, "Value is required." ) ;
      }

      if ( ! double.TryParse( value.ToString(), out var proposedValue ) ) {
        return new ValidationResult( false, "Value must be an integer value." ) ;
      }

      return proposedValue switch
      {
        < 0 => new ValidationResult( false, "Value must be greater than or equal to 0." ),
        > MAX_VALUE => new ValidationResult( false, $"Value must be less than or equal to {MAX_VALUE}." ),
        _ => ValidationResult.ValidResult
      } ;
    }
  }
}