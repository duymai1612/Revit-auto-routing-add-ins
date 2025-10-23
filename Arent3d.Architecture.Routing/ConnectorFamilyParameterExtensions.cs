using System ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public enum ConnectorFamilyType
  {
    Power,
    Sensor,
  }
  
  public static class ConnectorFamilyParameterExtensions
  {
    public static void SetConnectorFamilyType( this FamilyInstance connectorFamilyInstance, ConnectorFamilyType type )
    {
      connectorFamilyInstance.SetProperty( ConnectorFamilyParameter.ConnectorType, GetConnectorFamilyTypeName( type ) ) ;
    }

    public static ConnectorFamilyType? GetConnectorFamilyType( this FamilyInstance connectorFamilyInstance )
    {
      if ( false == connectorFamilyInstance.TryGetProperty( ConnectorFamilyParameter.ConnectorType, out string? typeName ) ) return null ;
      return GetConnectorFamilyTypeFromName( typeName! ) ;
    }

    private static string GetConnectorFamilyTypeName( ConnectorFamilyType type ) => type.ToString() ;

    private static ConnectorFamilyType? GetConnectorFamilyTypeFromName( string typeName )
    {
      if ( false == Enum.TryParse( typeName, true, out ConnectorFamilyType type ) ) return null ;
      return type ;
    }
  }
}