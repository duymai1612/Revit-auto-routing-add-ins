namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class PickUpModel
  {
    public string Item { get ; set ; }
    public string Floor { get ; set ; }
    public string ConstructionItems { get ; set ; }
    public string EquipmentType { get ; set ; }
    public string ProductName { get ; set ; }
    public string Use { get ; set ; }
    public string UsageName { get ; set ; }
    public string Construction { get ; set ; }
    public string ModelNumber { get ; set ; }
    public string Specification { get ; set ; }
    public string Specification2 { get ; set ; }
    public string Size { get ; set ; }
    public string Quantity { get ; set ; }
    public string Tani { get ; set ; }
    public string Supplement { get ; set ; }
    public string Supplement2 { get ; set ; }
    public string Group { get ; set ; }
    public string Layer { get ; set ; }
    public string Classification { get ; set ; }
    public string Standard { get ; set ; }
    public string PickUpNumber { get ; set ; }
    public string Direction { get ; set ; }
    public string ProductCode { get ; set ; }

    public PickUpModel( string? item, string? floor, string? constructionItems, string? equipmentType, string? productName, string? use, string? usageName, string? construction, string? modelNumber, string? specification, string? specification2, string? size, string? quantity, string? tani, string? supplement, string? supplement2, string? group, string? layer, string? classification, string? standard, string? pickUpNumber, string? direction, string? productCode )
    {
      Item = item ?? string.Empty ;
      Floor = floor ?? string.Empty ;
      ConstructionItems = constructionItems ?? string.Empty ;
      EquipmentType = equipmentType ?? string.Empty ;
      ProductName = productName ?? string.Empty ;
      Use = use ?? string.Empty ;
      UsageName = usageName ?? string.Empty ;
      Construction = construction ?? string.Empty ;
      ModelNumber = modelNumber ?? string.Empty ;
      Specification = specification ?? string.Empty ;
      Specification2 = specification2 ?? string.Empty ;
      Size = size ?? string.Empty ;
      Quantity = quantity ?? string.Empty ;
      Tani = tani ?? string.Empty ;
      Supplement = supplement ?? string.Empty ;
      Supplement2 = supplement2 ?? string.Empty ;
      Group = group ?? string.Empty ;
      Layer = layer ?? string.Empty ;
      Classification = classification ?? string.Empty ;
      Standard = standard ?? string.Empty ;
      PickUpNumber = pickUpNumber ?? string.Empty ;
      Direction = direction ?? string.Empty ;
      ProductCode = productCode ?? string.Empty ;
    }
  }
}