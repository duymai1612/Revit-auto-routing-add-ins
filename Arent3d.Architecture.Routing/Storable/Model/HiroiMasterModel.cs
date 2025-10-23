namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class HiroiMasterModel
  {
    public string Setubisyu { get ; set ; }
    public string Syurui { get ; set ; }
    public string Buzaicd { get ; set ; }
    public string Hinmei { get ; set ; }
    public string Kikaku { get ; set ; }
    public string Tani { get ; set ; }
    public string Buzaisyu { get ; set ; }
    public string Hinmeicd { get ; set ; }
    public string Ryakumeicd { get ; set ; }
    public string Type { get ; set ; }
    public string Size1 { get ; set ; }
    public string Size2 { get ; set ; }

    public HiroiMasterModel( string? setubisyu, string? syurui, string? buzaicd, string? hinmei, string? kikaku, string? tani, string? buzaisyu, string? hinmeicd, string? ryakumeicd, string? type, string? size1, string? size2 )
    {
      Setubisyu = setubisyu ?? string.Empty ;
      Syurui = syurui ?? string.Empty ;
      Buzaicd = buzaicd ?? string.Empty ;
      Hinmei = hinmei ?? string.Empty ;
      Kikaku = kikaku ?? string.Empty ;
      Tani = tani ?? string.Empty ;
      Buzaisyu = buzaisyu ?? string.Empty ;
      Hinmeicd = hinmeicd ?? string.Empty ;
      Ryakumeicd = ryakumeicd ?? string.Empty ;
      Type = type ?? string.Empty ;
      Size1 = size1 ?? string.Empty ;
      Size2 = size2 ?? string.Empty ;
    }
  }
}