using System ;
using System.Linq ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public enum RoutingParameter
  {
    // RoutingSharedParameters
    [ParameterGuid( "42a113b5-364a-4918-a423-6590c47b828f" ), NameOnRevit( "Route Name" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    RouteName,

    [ParameterGuid( "4620ee8d-7c76-4798-bfdc-87491ff8b355" ), NameOnRevit( "SubRoute Index" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    SubRouteIndex,

    [ParameterGuid( "58fd42f8-df12-41f3-9d7b-3dd4f1bffb41" ), NameOnRevit( "Branch Route Names" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    BranchRouteNames,
    
    [ParameterGuid( "b113fb98-a9eb-4f8e-a6a2-c4632922eb1b" ), NameOnRevit( "Route From-side Connector Ids" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    RoutedElementFromSideConnectorIds,

    [ParameterGuid( "6b594a61-ebec-4bc9-bbfb-e5abda7372cb" ), NameOnRevit( "Route To-side Connector Ids" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    RoutedElementToSideConnectorIds,

    [ParameterGuid( "5e822fe8-274e-41e0-b197-27a75bd52500" ), NameOnRevit( "Nearest From-side End Points" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    NearestFromSideEndPoints,

    [ParameterGuid( "7a6ec320-3c94-489a-89d2-dfb783ae8ae1" ), NameOnRevit( "Nearest To-side End Points" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    NearestToSideEndPoints,

    [ParameterGuid( "6cf2fece-a396-43e8-bede-4a2a75de5511" ), NameOnRevit( "Related Pass Point Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.PassPoints ) )]
    RelatedPassPointId,

    [ParameterGuid( "0e79cbf5-ac77-4fd2-be12-7969f5204a28" ), NameOnRevit( "Related Terminate Point Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.PassPoints ) )]
    RelatedTerminatePointId,
    
    [ParameterGuid( "ba87dfa3-c1f7-4b44-b07d-00333404bda8" ), NameOnRevit( "Representative Route Name" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    RepresentativeRouteName,
    
    [ParameterGuid( "9e825887-84fe-474d-ac2e-c683f7376647" ), NameOnRevit( "Representative SubRoute Index" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RoutingElements ) )]
    RepresentativeSubRouteIndex,

    [ParameterGuid("7632D393-DADE-437A-96A7-C4D508383012"), NameOnRevit("Rack Type")]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.RackTypeElements ) )]
    RackType,

    [ParameterGuid( "f208f9ab-b763-4b2a-afc9-0b2a22936dab" ), NameOnRevit( "Parent Envelope Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.CommonRoutingElement ) )]
    ParentEnvelopeId,

    [ParameterGuid( "f339149b-704c-403c-a97c-335646773992" ), NameOnRevit( "To-Side Connector Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.CableTrays ) )]
    ToSideConnectorId,

    [ParameterGuid( "57332190-02d7-4f25-a60d-b33a459f9fb7" ), NameOnRevit( "From-Side Connector Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.CableTrays ) )]
    FromSideConnectorId,
  }

  public enum PassPointParameter
  {
    [ParameterGuid( "b975f161-499f-4cc6-8e11-0d7ddf25b1f4" ), NameOnRevit( "PassPoint From-side Element Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.PassPoints ) )]
    PassPointNextToFromSideConnectorIds,

    [ParameterGuid( "7af4819d-3aec-4235-9f81-e6d3d0ca9ca2" ), NameOnRevit( "PassPoint To-side Element Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.PassPoints ) )]
    PassPointNextToToSideConnectorIds,

    [ParameterGuid( "c766d041-3867-4e55-a2bc-0272d8eb3013" ), NameOnRevit( "Related Connector Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.PassPoints ) )]
    RelatedConnectorId,

    [ParameterGuid( "86b66529-f4f6-4392-b80b-5f2dc71e9564" ), NameOnRevit( "Related From Connector Id" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.PassPoints ) )]
    RelatedFromConnectorId,
  }

  public enum RoutingFamilyLinkedParameter
  {
    [ParameterGuid( "3285f3e8-1838-4eba-a676-1a2af4708e7a" ), NameOnRevit( "Route Connector Relation Ids" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.ElementsUsedForUI ) )]
    RouteConnectorRelationIds,

    [ParameterGuid( "f054f110-68e7-4bce-9d17-688557d410da" ), NameOnRevit( "Construction Item" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.ConstructionItems ) )]
    ConstructionItem
  }

  public enum ConnectorFamilyParameter
  {
    [ParameterGuid( "442b05ee-df38-4595-93c9-e2d7cfa227e9" ), NameOnRevit( "Connector Type" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.Connectors ) )]
    ConnectorType,
            
    [ParameterGuid( "f71cbd72-3fe4-47cb-b777-36d6511d42ed" ), NameOnRevit( "CeeD Code" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.PG_IDENTITY_DATA, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.Connectors ) )]
    CeeDCode,
  }

  public enum BranchNumberParameter
  {
    [ParameterGuid( "01c73735-4b79-4729-91af-3dede453c482" ), NameOnRevit( "BranchNumber" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.INVALID, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.SpaceElements ) )]
    BranchNumber
  }
  
  public enum AHUNumberParameter
  {
    [ParameterGuid( "e1522ec9-2d76-4c32-addc-f9f27d3aa8ea" ), NameOnRevit( "AHUNumber" )]
    [BuiltInCategories( ExternalParameterType.Instance, BuiltInParameterGroup.INVALID, typeof( BuiltInCategorySets ), nameof( BuiltInCategorySets.AHUNumberElements ) )]
    AHUNumber
  }
  

  public static class RoutingPropertyExtensions
  {
    public static bool AllRoutingParametersAreRegistered( this Document document )
    {
      return document.AllParametersAreRegistered<RoutingParameter>() && document.AllParametersAreRegistered<RoutingFamilyLinkedParameter>() ;
    }

    public static void MakeCertainAllRoutingParameters( this Document document )
    {
      document.LoadAllParametersFromFile<RoutingParameter>( AssetManager.GetRoutingSharedParameterPath() ) ;
      document.LoadAllParametersFromFile<PassPointParameter>( AssetManager.GetPassPointSharedParameterPath() ) ;
      document.LoadAllParametersFromFile<RoutingFamilyLinkedParameter>( AssetManager.GetRoutingElementSharedParameterPath() );
      document.LoadAllParametersFromFile<ConnectorFamilyParameter>( AssetManager.GetConnectorSharedParameterPath() ) ;
    }

    public static void MakeBranchNumberParameter( this Document document )
    {
      document.LoadAllParametersFromFile<BranchNumberParameter>( AssetManager.GetSpaceSharedParameterPath() ) ;
    }

    public static void UnloadAllRoutingParameters( this Document document )
    {
      document.UnloadAllParametersFromFile<RoutingParameter>( AssetManager.GetRoutingSharedParameterPath() ) ;
      document.UnloadAllParametersFromFile<PassPointParameter>( AssetManager.GetPassPointSharedParameterPath() ) ;
      document.UnloadAllParametersFromFile<RoutingFamilyLinkedParameter>( AssetManager.GetRoutingElementSharedParameterPath() ) ;
      document.UnloadAllParametersFromFile<ConnectorFamilyParameter>( AssetManager.GetConnectorSharedParameterPath() ) ;
      document.UnloadAllParametersFromFile<BranchNumberParameter>( AssetManager.GetSpaceSharedParameterPath() ) ;
    }
    
    public static void MakeAHUNumberParameter( this Document document )
    {
      document.LoadAllParametersFromFile<AHUNumberParameter>(AssetManager.GetSpaceSharedParameterPath()) ;
    }
  }
}
