using System ;
using System.Globalization ;
using System.Text ;
using System.Windows ;
using System.Windows.Data ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Forms.ValueConverters
{
  public class ShaftToStringConverter : IMultiValueConverter
  {
    public static IMultiValueConverter Instance { get ; } = new ShaftToStringConverter() ;

    public object Convert( object[] values, Type targetType, object parameter, CultureInfo culture )
    {
      if ( 2 != values.Length ) return DependencyProperty.UnsetValue ;

      if ( values[ 0 ] is not Opening opening ) return "Dialog.Electrical.Shaft.List.None".GetAppStringByKeyOrDefault( "(None)" ) ;

      return values[ 1 ] switch
      {
        DisplayUnit.METRIC => AsMetric( opening ),
        DisplayUnit.IMPERIAL => AsImperial( opening ),
        _ => AsImperial( opening ),
      } ;
    }

    public object[] ConvertBack( object value, Type[] targetTypes, object parameter, CultureInfo culture ) => throw new NotSupportedException() ;

    private static string AsMetric( Opening opening )
    {
      var topLevel = opening.get_Parameter( BuiltInParameter.WALL_HEIGHT_TYPE )?.AsValueString() ?? string.Empty ;
      var bottomLevel = opening.get_Parameter( BuiltInParameter.WALL_BASE_CONSTRAINT )?.AsValueString() ?? string.Empty ;

      var (x, y, z) = opening.GetShaftPosition() ;
      return $"{opening.Name}: {bottomLevel}-{topLevel} (X: {x.RevitUnitsToMeters():0.000} m, Y: {y.RevitUnitsToMeters():0.000} m, Z: {z.RevitUnitsToMeters():0.000} m)" ;
    }

    private static string AsImperial( Opening opening )
    {
      var topLevel = opening.get_Parameter( BuiltInParameter.WALL_HEIGHT_TYPE )?.AsValueString() ?? string.Empty ;
      var bottomLevel = opening.get_Parameter( BuiltInParameter.WALL_BASE_CONSTRAINT )?.AsValueString() ?? string.Empty ;

      var (x, y, z) = opening.GetShaftPosition() ;
      return $"{opening.Name}: {bottomLevel}-{topLevel} (X: {FeetAndInches(x.RevitUnitsToFeet())}, Y: {FeetAndInches(y.RevitUnitsToFeet())}, Z: {FeetAndInches(z.RevitUnitsToFeet())})" ;
    }

    private static string FeetAndInches( double feet )
    {
      if ( double.IsNaN( feet ) ) return "NaN" ;
      
      var builder = new StringBuilder() ;

      if ( feet < 0 ) {
        builder.Append( '-' ) ;
        feet = -feet ;
      }

      if ( double.IsInfinity( feet ) ) {
        builder.Append( "∞" ) ;
        return builder.ToString() ;
      }

      var inches = (int)( feet + 1.0 / 24.0 ) ;

      if ( 12 <= inches ) {
        builder.Append( inches / 12 ) ;
        builder.Append( '\'' ) ;
      }

      var inchValues = inches % 12 ;
      if ( 0 != inchValues ) {
        builder.Append( inchValues ) ;
        builder.Append( '\"' ) ;
      }

      return builder.ToString() ;
    }
  }
}