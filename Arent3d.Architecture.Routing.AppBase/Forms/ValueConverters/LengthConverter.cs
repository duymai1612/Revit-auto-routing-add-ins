using System ;
using System.Globalization ;
using System.Windows.Data ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing.AppBase.Forms.ValueConverters
{
  public abstract class LengthConverter : IValueConverter
  {
    public static LengthConverter Millimeters { get ; } = new MillimetersLengthConverter() ;
    public static LengthConverter Inches { get ; } = new InchesLengthConverter() ;
    public static LengthConverter Default => Millimeters ;

    public object Convert( object value, Type targetType, object parameter, CultureInfo culture )
    {
      if ( value is not double revitUnitLength ) return double.NaN ;
      return ConvertUnit( revitUnitLength ) ;
    }

    public object ConvertBack( object value, Type targetType, object parameter, CultureInfo culture )
    {
      if ( value is not double displayUnitLength ) return double.NaN ;
      return ConvertBackUnit( displayUnitLength ) ;
    }

    public abstract double ConvertUnit( double revitUnitLength ) ;
    public abstract double ConvertBackUnit( double displayUnitLength ) ;

    private class MillimetersLengthConverter : LengthConverter
    {
      public override double ConvertUnit( double revitUnitLength ) => revitUnitLength.RevitUnitsToMillimeters() ;
      public override double ConvertBackUnit( double displayUnitLength ) => displayUnitLength.MillimetersToRevitUnits() ;
    }

    private class InchesLengthConverter : LengthConverter
    {
      public override double ConvertUnit( double revitUnitLength ) => UnitUtils.ConvertFromInternalUnits( revitUnitLength, DisplayUnitTypes.Inches ) ;
      public override double ConvertBackUnit( double displayUnitLength ) => UnitUtils.ConvertToInternalUnits( displayUnitLength, DisplayUnitTypes.Inches ) ;
    }
  }
}