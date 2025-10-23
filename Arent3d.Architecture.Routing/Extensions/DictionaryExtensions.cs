using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Arent3d.Architecture.Routing.Extensions
{
  public static class DictionaryExtensions
  {

    public static TValue GetOrDefault<TKey, TValue>( this Dictionary<TKey, TValue> keyValues, TKey key, Func<TValue> defaultValue ) where TValue : class
    {
      return ( keyValues.TryGetValue( key, out var value ) ? value : null ) ?? defaultValue() ;
    }
  }
}
