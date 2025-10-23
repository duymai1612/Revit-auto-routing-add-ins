using System ;
using System.Linq ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public static class BuiltInCategorySets
  {
    public static readonly BuiltInCategory[] RackTypeElements =
    {
      BuiltInCategory.OST_CableTray,
      BuiltInCategory.OST_CableTrayFitting,
    };

    public static readonly BuiltInCategory[] Pipes =
    {
      BuiltInCategory.OST_FlexPipeCurves,
      BuiltInCategory.OST_PipeFitting,
      BuiltInCategory.OST_PipeCurves,
    } ;

    public static readonly BuiltInCategory[] Ducts =
    {
      BuiltInCategory.OST_DuctTerminal,
      BuiltInCategory.OST_DuctFitting,
      BuiltInCategory.OST_DuctCurves,
      BuiltInCategory.OST_PlaceHolderDucts,
      BuiltInCategory.OST_FlexDuctCurves,
    } ;

    public static readonly BuiltInCategory[] MechanicalRoutingElements = CombineArrays(
      Pipes,
      Ducts
    ) ;

    public static readonly BuiltInCategory[] OtherMechanicalElements =
    {
      BuiltInCategory.OST_DuctAccessory,
      BuiltInCategory.OST_DuctSystem,
      BuiltInCategory.OST_MechanicalEquipment,
      BuiltInCategory.OST_PipeAccessory,
      //BuiltInCategory.OST_PipeSegments, // cannot use parameters for OST_PipeSegments category!
      BuiltInCategory.OST_PlumbingFixtures,
      BuiltInCategory.OST_Sprinklers,
    } ;

    public static readonly BuiltInCategory[] Conduits =
    {
      BuiltInCategory.OST_Conduit,
      BuiltInCategory.OST_ConduitFitting,
      BuiltInCategory.OST_ConduitRun,
    } ;
    
    public static readonly BuiltInCategory[] ConstructionItems =
    {
      BuiltInCategory.OST_Conduit,
      BuiltInCategory.OST_ConduitFitting,
      BuiltInCategory.OST_ConduitRun,
      BuiltInCategory.OST_ElectricalFixtures,
      BuiltInCategory.OST_CableTray,
      BuiltInCategory.OST_CableTrayFitting
    } ;

    public static readonly BuiltInCategory[] CableTrays =
    {
      BuiltInCategory.OST_CableTray,
      BuiltInCategory.OST_CableTrayFitting,
    } ;

    public static readonly BuiltInCategory[] ElectricalRoutingElements = CombineArrays(
      Conduits,
      CableTrays
    ) ;

    public static readonly BuiltInCategory[] OtherElectricalElements =
    {
      BuiltInCategory.OST_ElectricalEquipment,
      BuiltInCategory.OST_ElectricalFixtures,
    } ;

    public static readonly BuiltInCategory[] CommonRoutingElement =
    {
      BuiltInCategory.OST_GenericModel,
    } ;

    public static readonly BuiltInCategory[] RoutingElements = CombineArrays(
      MechanicalRoutingElements,
      OtherMechanicalElements,
      ElectricalRoutingElements,
      OtherElectricalElements,
      CommonRoutingElement
    ) ;

    public static readonly BuiltInCategory[] Obstacles = CombineArrays(
      RoutingElements,
      new[]
      {
        BuiltInCategory.OST_StructuralColumns,
        BuiltInCategory.OST_StructuralFoundation,
        BuiltInCategory.OST_StructuralFraming,
        BuiltInCategory.OST_StructuralTruss,
      }
    ) ;

    public static readonly BuiltInCategory[] PassPoints =
    {
      BuiltInCategory.OST_MechanicalEquipment,
    } ;

    public static readonly BuiltInCategory[] ElementsUsedForUI = 
    {
      BuiltInCategory.OST_MechanicalEquipment,
      BuiltInCategory.OST_GenericModel,
    } ;
    
    public static readonly BuiltInCategory[] Connectors =
    {
      BuiltInCategory.OST_ElectricalFixtures,
    } ;

    public static readonly BuiltInCategory[] Fittings =
    {
      BuiltInCategory.OST_DuctFitting,
      BuiltInCategory.OST_PipeFitting,
      BuiltInCategory.OST_CableTrayFitting,
      BuiltInCategory.OST_ConduitFitting,
    } ;

    public static readonly BuiltInCategory[] CenterLineCategories =
    {
      BuiltInCategory.OST_CenterLines,
      BuiltInCategory.OST_DuctCurvesCenterLine,
      BuiltInCategory.OST_DuctFittingCenterLine,
      BuiltInCategory.OST_FlexDuctCurvesCenterLine,
      BuiltInCategory.OST_PipeCurvesCenterLine,
      BuiltInCategory.OST_PipeFittingCenterLine,
      BuiltInCategory.OST_FlexPipeCurvesCenterLine,
    } ;
    
    public static readonly BuiltInCategory[] SpaceElements =
    {
      BuiltInCategory.OST_MEPSpaces
    } ;
    
    public static readonly BuiltInCategory[] PickUpElements =
    {
      BuiltInCategory.OST_ElectricalEquipment,
      BuiltInCategory.OST_ElectricalFixtures,
      BuiltInCategory.OST_MechanicalEquipment,
    } ;

    public static readonly BuiltInCategory[] AHUNumberElements =
    {
      BuiltInCategory.OST_MEPSpaces,
      BuiltInCategory.OST_MechanicalEquipment,
    } ;
    
    private static T[] CombineArrays<T>( params T[][] arrays )
    {
      var totalLength = arrays.Sum( array => array.Length ) ;
      var result = new T[ totalLength ] ;
      var index = 0 ;
      foreach ( var array in arrays ) {
        Array.Copy( array, 0, result, index, array.Length ) ;
        index += array.Length ;
      }
      return result ;
    }
  }
}