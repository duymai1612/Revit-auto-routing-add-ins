using System.IO ;
using System.Reflection ;

namespace Arent3d.Architecture.Routing
{
  public static class AssetManager
  {
#if REVIT2019
    private const string FamilyFolderName = @"Families\2019" ;
#elif REVIT2020
    private const string FamilyFolderName = @"Families\2020" ;
#elif REVIT2021
    private const string FamilyFolderName = @"Families\2021" ;
#elif REVIT2022
    private const string FamilyFolderName = @"Families\2022" ;
#endif
    private const string SettingFolderName = "SharedParameterFile" ;
    private const string CsvFolderName = "CsvFiles" ;

    private const string RoutingSharedParameterFileName = "RoutingSharedParameters.txt" ;
    private const string PassPointSharedParameterFileName = "PassPointSharedParameters.txt" ;
    private const string RoutingElementSharedParameterFileName = "RoutingElementSharedParameters.txt";
    private const string ConnectorSharedParameterFileName = "ConnectorSharedParameters.txt" ;
    private const string SpaceSharedParameterFileName = "SpaceSharedParameters.txt" ;

    private static readonly string AssetPath = Path.Combine( Path.GetDirectoryName( Assembly.GetExecutingAssembly().Location )!, "Assets" ) ;

    public static string GetFamilyPath( string familyName )
    {
      return GetPath( FamilyFolderName, familyName + ".rfa" ) ;
    }

    public static string GetRoutingSharedParameterPath()
    {
      return GetPath( SettingFolderName, RoutingSharedParameterFileName ) ;
    }

    public static string GetPassPointSharedParameterPath()
    {
      return GetPath( SettingFolderName, PassPointSharedParameterFileName ) ;
    }

    public static string GetRoutingElementSharedParameterPath()
    {
        return GetPath( SettingFolderName, RoutingElementSharedParameterFileName );
    }
    public static string GetConnectorSharedParameterPath()
    {
      return GetPath( SettingFolderName, ConnectorSharedParameterFileName );
    }

    public static string GetSpaceSharedParameterPath()
    {
      return GetPath( SettingFolderName, SpaceSharedParameterFileName );
    }

    public static string GetCeeDModelPath( string ceeDFileName )
    {
      return GetPath( CsvFolderName, ceeDFileName + ".xlsx" ) ;
    }
    
    private static string GetPath( string folderName, string fileName )
    {
      return Path.Combine( AssetPath, folderName, fileName ) ;
    }
  }
}