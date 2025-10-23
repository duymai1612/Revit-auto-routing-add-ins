using System ;
using System.Collections.Generic ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// Manages current <see cref="Document"/> of Revit.
  /// </summary>
  public static class DocumentMapper
  {
    private static readonly Dictionary<Document, DocumentData> _mapper = new() ;

    public static DocumentData Get( Document document )
    {
      return _mapper.TryGetValue( document, out var data ) ? data : throw new KeyNotFoundException() ;
    }

    public static void Register( Document document )
    {
      if ( _mapper.ContainsKey( document ) ) return ; // duplicated

      _mapper.Add( document, new DocumentData( document ) ) ;
      
      // TODO: search auto routing families
    }

    public static void Unregister( Document document )
    {
      _mapper.Remove( document ) ;
    }
  }
}