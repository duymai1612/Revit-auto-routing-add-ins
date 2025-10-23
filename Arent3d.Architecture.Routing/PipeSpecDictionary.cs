using System ;
using System.Collections ;
using System.Collections.Generic ;
using Arent3d.Architecture.Routing.FittingSizeCalculators ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  internal sealed class RouteMEPSystemKey : IEquatable<RouteMEPSystemKey>
  {
    private readonly int _mepSystemTypeId ;
    private readonly int _mepCurveTypeId ;

    public RouteMEPSystemKey( SubRoute subRoute )
    {
      _mepSystemTypeId = ( subRoute.Route.GetMEPSystemType()?.Id ?? ElementId.InvalidElementId ).IntegerValue ;
      _mepCurveTypeId = subRoute.GetMEPCurveType().Id.IntegerValue ;
    }

    public bool Equals( RouteMEPSystemKey other )
    {
      return ( this._mepSystemTypeId == other._mepSystemTypeId && this._mepCurveTypeId == other._mepCurveTypeId ) ;
    }

    public override bool Equals( object? obj )
    {
      if ( ReferenceEquals( null, obj ) ) return false ;
      if ( ReferenceEquals( this, obj ) ) return true ;

      return Equals( (RouteMEPSystemKey)obj ) ;
    }

    public override int GetHashCode()
    {
      unchecked {
        return ( _mepSystemTypeId * 397 ) ^ _mepCurveTypeId ;
      }
    }

    public static bool operator ==( RouteMEPSystemKey? left, RouteMEPSystemKey? right )
    {
      return Equals( left, right ) ;
    }

    public static bool operator !=( RouteMEPSystemKey? left, RouteMEPSystemKey? right )
    {
      return ! Equals( left, right ) ;
    }
  }

  public class PipeSpecDictionary
  {
    private readonly Document _document ;
    private readonly IFittingSizeCalculator _fittingSizeCalculator ;
    private readonly Dictionary<RouteMEPSystemKey, (RouteMEPSystem RouteMEPSystem, MEPSystemPipeSpec PipeSpec)> _dic = new() ;

    public PipeSpecDictionary( Document document, IFittingSizeCalculator fittingSizeCalculator )
    {
      _document = document ;
      _fittingSizeCalculator = fittingSizeCalculator ;
    }

    public RouteMEPSystem GetRouteMEPSystem( SubRoute subRoute ) => GetTuple( subRoute ).RouteMEPSystem ;
    public MEPSystemPipeSpec GetMEPSystemPipeSpec( SubRoute subRoute ) => GetTuple( subRoute ).PipeSpec ;

    private (RouteMEPSystem RouteMEPSystem, MEPSystemPipeSpec PipeSpec) GetTuple( SubRoute subRoute )
    {
      var key = new RouteMEPSystemKey( subRoute ) ;

      if ( false == _dic.TryGetValue( key, out var tuple ) ) {
        var mepSystem = new RouteMEPSystem( _document, subRoute ) ;
        var pipeSpec = new MEPSystemPipeSpec( mepSystem, _fittingSizeCalculator ) ;
        tuple = ( mepSystem, pipeSpec ) ;
        _dic.Add( key, tuple ) ;
      }

      return tuple ;
    }
  }
}