using System ;
using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Structure ;

namespace Arent3d.Architecture.Routing
{
  public enum RoutingFamilyType
  {
    [NameOnRevit( "Routing Rack Guide" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    RackGuide,

    [NameOnRevit( "Routing Rack Space" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    RackSpace,

    [NameOnRevit( "Routing Shaft" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    Shaft,

    [NameOnRevit( "Routing Pass Point" )]
    [FamilyCategory( BuiltInCategory.OST_MechanicalEquipment )]
    PassPoint,

    [NameOnRevit( "Routing Terminate Point" )]
    [FamilyCategory( BuiltInCategory.OST_MechanicalEquipment )]
    TerminatePoint,

    [NameOnRevit( "Routing Corn Point" )]
    [FamilyCategory( BuiltInCategory.OST_MechanicalEquipment )]
    CornPoint,

    [NameOnRevit( "Routing Connector Point" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    ConnectorPoint,

    [NameOnRevit( "Routing Connector In Point" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    ConnectorInPoint,

    [NameOnRevit( "Routing Connector Out Point" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    ConnectorOutPoint,

    [NameOnRevit("電線管用ファミリ_ver1.0")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalFixtures)]
    ConnectorTwoSide,
    
    [NameOnRevit("電線管用ファミリ(片側のみ)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalFixtures)]
    ConnectorOneSide,

    [NameOnRevit("ダクト用湿度ｾﾝｻｰ(ロゴあり)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    HumiditySensorForDuctWithLogo,

    [NameOnRevit("ダクト用湿度ｾﾝｻｰ(ロゴなし)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    HumiditySensorForDuctWithoutLogo,

    [NameOnRevit("ダンパ操作器")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    DamperActuator,

    [NameOnRevit("室内用湿度ｾﾝｻｰ(ロゴあり)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    IndoorHumiditySensorWithLogo,

    [NameOnRevit("室内用湿度ｾﾝｻｰ(ロゴなし)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    IndoorHumiditySensorWithoutLogo,

    [NameOnRevit("電動二方弁(ロゴあり)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    ElectricTwoWayValveWithLogo,

    [NameOnRevit("電動二方弁(ロゴなし)")]
    [FamilyCategory(BuiltInCategory.OST_ElectricalEquipment)]
    ElectricTwoWayValveWithoutLogo,

    [NameOnRevit( "Routing Envelope" )]
    [FamilyCategory( BuiltInCategory.OST_GenericModel )]
    Envelope,
    
    [NameOnRevit( "Cable Tray" )]
    [FamilyCategory( BuiltInCategory.OST_CableTrayFitting )]
    CableTray,
    
    [NameOnRevit( "Cable Tray Elbow" )]
    [FamilyCategory( BuiltInCategory.OST_CableTrayFitting )]
    CableTrayFitting,

    [NameOnRevit("Cable Tray Reducer")]
    [FamilyCategory(BuiltInCategory.OST_CableTrayFitting)]
    CableTrayReducer,

    [NameOnRevit( "SA_FASU(F4-150 200Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F4_150_200Phi,
    
    [NameOnRevit( "SA_FASU(F4-150 250Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F4_150_250Phi,
    
    [NameOnRevit( "SA_FASU(F5-150 250Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F5_150_250Phi,
    
    [NameOnRevit( "SA_FASU(F6-150 250Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F6_150_250Phi,
    
    [NameOnRevit( "SA_FASU(F6-150 300Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F6_150_300Phi,
    
    [NameOnRevit( "SA_FASU(F7-150 300Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F7_150_300Phi,
    
    [NameOnRevit( "SA_FASU(F8-150 250Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F8_150_250Phi,
    
    [NameOnRevit( "SA_FASU(F8-150 300Φ)" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    FASU_F8_150_300Phi,

    [NameOnRevit( "SA_VAV" )]
    [FamilyCategory( BuiltInCategory.OST_DuctAccessory)]
    TTE_VAV_140,
    
    [NameOnRevit( "Direction Cylindrical Shaft" )]
    [FamilyCategory( BuiltInCategory.OST_DetailComponents)]
    DirectionCylindricalShaft,
  }

  public static class RoutingFamilyExtensions
  {
    public static bool AllRoutingFamiliesAreLoaded( this Document document ) => document.AllFamiliesAreLoaded<RoutingFamilyType>() ;

    public static void MakeCertainAllRoutingFamilies( this Document document ) => document.MakeCertainAllFamilies<RoutingFamilyType>( AssetManager.GetFamilyPath ) ;
    public static void EraseAllRoutingFamilies( this Document document ) => document.UnloadAllFamilies<RoutingFamilyType>() ;

    public static FamilyInstance Instantiate( this FamilySymbol symbol, XYZ position, string levelName, StructuralType structuralType )
    {
      var level = GetLevel( symbol.Document, levelName ) ;
      if ( null == level ) throw new InvalidOperationException() ;
      return symbol.Instantiate( position, level, structuralType ) ;
    }
    public static FamilyInstance Instantiate( this FamilySymbol symbol, XYZ position, Level level, StructuralType structuralType )
    {
      var document = symbol.Document ;
      if ( false == symbol.IsActive ) symbol.Activate() ;

      return document.Create.NewFamilyInstance( position, symbol, level, structuralType ) ;
    }

    private static Level? GetLevel( Document document, string levelName )
    {
      return document.GetAllElements<Level>().FirstOrDefault( l => l.Name == levelName ) ;
    }
  }
}