namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class HiroiSetMasterModel
  {
    public string ParentPartModelNumber { get ; set ; }
    public string ParentPartName { get ; set ; }
    public string ParentPartsQuantity { get ; set ; }
    public string MaterialCode1 { get ; set ; }
    public string Name1 { get ; set ; }
    public string Quantity1 { get ; set ; }
    public string MaterialCode2 { get ; set ; }
    public string Name2 { get ; set ; }
    public string Quantity2 { get ; set ; }
    public string MaterialCode3 { get ; set ; }
    public string Name3 { get ; set ; }
    public string Quantity3 { get ; set ; }
    public string MaterialCode4 { get ; set ; }
    public string Name4 { get ; set ; }
    public string Quantity4 { get ; set ; }
    public string MaterialCode5 { get ; set ; }
    public string Name5 { get ; set ; }
    public string Quantity5 { get ; set ; }
    public string MaterialCode6 { get ; set ; }
    public string Name6 { get ; set ; }
    public string Quantity6 { get ; set ; }
    public string MaterialCode7 { get ; set ; }
    public string Name7 { get ; set ; }
    public string Quantity7 { get ; set ; }
    public string MaterialCode8 { get ; set ; }
    public string Name8 { get ; set ; }
    public string Quantity8 { get ; set ; }

    public HiroiSetMasterModel( string? parentPartModelNumber, string? parentPartName, string? parentPartsQuantity, string? materialCode1, string? name1, string? quantity1, string? materialCode2, string? name2, string? quantity2, string? materialCode3, string? name3, string? quantity3, string? materialCode4, string? name4, string? quantity4, string? materialCode5, string? name5, string? quantity5, string? materialCode6, string? name6, string? quantity6, string? materialCode7, string? name7, string? quantity7, string? materialCode8, string? name8, string? quantity8 )
    {
      ParentPartModelNumber = parentPartModelNumber ?? string.Empty ;
      ParentPartName = parentPartName ?? string.Empty ;
      ParentPartsQuantity = parentPartsQuantity ?? string.Empty ;
      MaterialCode1 = materialCode1 ?? string.Empty ;
      Name1 = name1 ?? string.Empty ;
      Quantity1 = quantity1 ?? string.Empty ;
      MaterialCode2 = materialCode2 ?? string.Empty ;
      Name2 = name2 ?? string.Empty ;
      Quantity2 = quantity2 ?? string.Empty ;
      MaterialCode3 = materialCode3 ?? string.Empty ;
      Name3 = name3 ?? string.Empty ;
      Quantity3 = quantity3 ?? string.Empty ;
      MaterialCode4 = materialCode4 ?? string.Empty ;
      Name4 = name4 ?? string.Empty ;
      Quantity4 = quantity4 ?? string.Empty ;
      MaterialCode5 = materialCode5 ?? string.Empty ;
      Name5 = name5 ?? string.Empty ;
      Quantity5 = quantity5 ?? string.Empty ;
      MaterialCode6 = materialCode6 ?? string.Empty ;
      Name6 = name6 ?? string.Empty ;
      Quantity6 = quantity6 ?? string.Empty ;
      MaterialCode7 = materialCode7 ?? string.Empty ;
      Name7 = name7 ?? string.Empty ;
      Quantity7 = quantity7 ?? string.Empty ;
      MaterialCode8 = materialCode8 ?? string.Empty ;
      Name8 = name8 ?? string.Empty ;
      Quantity8 = quantity8 ?? string.Empty ;
    }
  }
}