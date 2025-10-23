using System ;
using System.Collections.Generic ;

namespace Arent3d.Architecture.Routing
{
  internal class SizeTable<TValue, TResult> where TValue : struct where TResult : struct
  {
    private readonly Func<TValue, TResult> _generator ;
    private readonly Dictionary<TValue, TResult> _dic = new() ;

    public SizeTable( Func<TValue, TResult> generator )
    {
      _generator = generator ;
    }

    public TResult Get( TValue value )
    {
      if ( false == _dic.TryGetValue( value, out var result ) ) {
        result = _generator( value ) ;
        _dic.Add( value, result ) ;
      }

      return result ;
    }
  }
}