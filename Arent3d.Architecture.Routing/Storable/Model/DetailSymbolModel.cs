namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class DetailSymbolModel
  {
    public string DetailSymbolId { get ; set ; }
    public string DetailSymbol { get ; set ; }
    public string ConduitId { get ; set ; }
    public string RouteName { get ; set ; }
    public string Code { get ; set ; }
    public string LineIds { get ; set ; }
    public bool IsParentSymbol { get ; set ; }

    public DetailSymbolModel( string? detailSymbolId, string? detailSymbol, string? conduitId, string? routeName, string? code, string? lineIds, bool? isParentSymbol )
    {
      DetailSymbolId = detailSymbolId ?? string.Empty ;
      DetailSymbol = detailSymbol ?? string.Empty ;
      ConduitId = conduitId ?? string.Empty ;
      RouteName = routeName ?? string.Empty ;
      Code = code ?? string.Empty ;
      LineIds = lineIds ?? string.Empty ;
      IsParentSymbol = isParentSymbol ?? true ;
    }
  }
}