using System ;

namespace Arent3d.Architecture.Routing
{
  public enum FixedHeightType
  {
    Floor,
    Ceiling,
  }

  public readonly struct FixedHeight
  {
    public static FixedHeight? CreateOrNull( string? typeName, double? height )
    {
      if ( false == Enum.TryParse( typeName, out FixedHeightType type ) ) return null ;
      if ( height is not { } nonNullHeight ) return null ;
      return new FixedHeight( type, nonNullHeight ) ;
    }

    public static FixedHeight? CreateOrNull( FixedHeightType? type, double? height )
    {
      if ( type is not { } nonNullType || height is not { } nonNullHeight ) return null ;
      return new FixedHeight( nonNullType, nonNullHeight ) ;
    }

    private FixedHeight( FixedHeightType type, double height )
    {
      Type = type ;
      Height = height ;
    }

    public FixedHeightType Type { get ; }
    public double Height { get ; }
  }
}