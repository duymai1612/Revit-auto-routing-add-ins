namespace Arent3d.Architecture.Routing.Storable.Model
{
  public class OffsetSettingModel
  {
    private const double DEFAULT_OFFSET = 600 ;
    public double Offset { get ; set ; }

    public OffsetSettingModel()
    {
      Offset = DEFAULT_OFFSET ;
    }

    public OffsetSettingModel( double? offset )
    {
      Offset = offset ?? 0 ;
    }
  }
}