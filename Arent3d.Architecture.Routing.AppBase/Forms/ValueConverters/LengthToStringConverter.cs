using System ;
using System.Globalization ;
using System.Windows.Data ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Forms.ValueConverters
{
  public abstract class LengthToStringConverter : IValueConverter
  {
    public static LengthToStringConverter Millimeters { get ; } = new MillimetersLengthToStringConverter() ;
    public static LengthToStringConverter Inches { get ; } = new InchesLengthToStringConverter() ;
    public static LengthToStringConverter Default => Millimeters ;

    public object Convert( object value, Type targetType, object parameter, CultureInfo culture )
    {
      if ( value is not double diameter ) return string.Empty ;
      return ToDisplayString( diameter, culture ) ;
    }

    public object ConvertBack( object value, Type targetType, object parameter, CultureInfo culture ) => throw new NotSupportedException() ;

    protected abstract string ToDisplayString( double diameter, CultureInfo culture ) ;

    private class MillimetersLengthToStringConverter : LengthToStringConverter
    {
      protected override string ToDisplayString( double diameter, CultureInfo culture )
      {
        var millimeters = diameter.RevitUnitsToMillimeters() ;
        return $"{Math.Round( millimeters, 2, MidpointRounding.AwayFromZero ).ToString( culture )}mm" ;
      }
    }

    private class InchesLengthToStringConverter : LengthToStringConverter
    {
      protected override string ToDisplayString( double diameter, CultureInfo culture )
      {
        var inches = UnitUtils.ConvertFromInternalUnits( diameter, DisplayUnitTypes.Inches ) ;
        return $"{Math.Round( inches, 2, MidpointRounding.AwayFromZero ).ToString( culture )}\"" ;
      }
    }
  }
}