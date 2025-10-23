namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class HiroiSetCdMasterModel
  {
    public string SetCode { get ; set ; }
    public string QuantityParentPartModelNumber { get ; set ; }
    public string LengthParentPartModelNumber { get ; set ; }
    public string ConstructionClassification { get ; set ; }

    public HiroiSetCdMasterModel( string? setCode, string? quantityParentPartModelNumber, string? lengthParentPartModelNumber, string? constructionClassification )
    {
      SetCode = setCode ?? string.Empty ;
      QuantityParentPartModelNumber = quantityParentPartModelNumber ?? string.Empty ;
      LengthParentPartModelNumber = lengthParentPartModelNumber ?? string.Empty ;
      ConstructionClassification = constructionClassification ?? string.Empty ;
    }
  }
}