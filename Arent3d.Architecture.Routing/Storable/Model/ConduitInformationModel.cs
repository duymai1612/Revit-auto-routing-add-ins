using System.ComponentModel ;

namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class ConduitInformationModel
  {
    public bool? CalculationExclusion { get ; set ; }
    public string? Floor { get ; set ; }
    public string? CeeDCode { get ; set ; }
    public string? DetailSymbol { get ; set ; }
    public string? WireType { get ; set ; }
    public string? WireSize { get ; set ; }
    public string? WireStrip { get ; set ; }
    public string? WireBook { get ; set ; }
    public string? EarthType { get ; set ; }
    public string? EarthSize { get ; set ; }
    public string? NumberOfGrounds { get ; set ; }
    public string? PipingType { get ; set ; }
    public string? PipingSize { get ; set ; }
    public string? NumberOfPipes { get ; set ; }
    public string? ConstructionClassification { get ; set ; }
    public string? Classification { get ; set ; }
    public string? ConstructionItems { get ; set ; }
    public string? PlumbingItems { get ; set ; }
    public string? Remark { get ; set ; }

    public  int Quantity { get ; set ; }

    public ConduitInformationModel( 
      bool? calculationExclusion, 
      string? floor, 
      string? ceeDCode,
      string? detailSymbol,
      string? wireType, 
      string? wireSize, 
      string? wireStrip, 
      string? wireBook,
      string? earthType,
      string? earthSize,
      string? numberOfGrounds,
      string? pipingType,
      string? pipingSize,
      string? numberOfPipes,
      string? constructionClassification,
      string? classification,
      string? constructionItems,
      string? plumbingItems,
      string? remark)
    {
      CalculationExclusion = calculationExclusion ;
      Floor = floor ;
      CeeDCode = ceeDCode ;
      DetailSymbol = detailSymbol ;
      WireType = wireType ;
      WireSize = wireSize ;
      WireStrip = wireStrip ;
      WireBook = wireBook ;
      EarthType = earthType ;
      EarthSize = earthSize ;
      NumberOfGrounds = numberOfGrounds ;
      PipingType = pipingType ;
      PipingSize = pipingSize ;
      NumberOfPipes = numberOfPipes ;
      ConstructionClassification = constructionClassification ;
      Classification = classification ;
      ConstructionItems = constructionItems ;
      PlumbingItems = plumbingItems ;
      Remark = remark ;
    }
  }
}