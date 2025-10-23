using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Reflection ;
using Arent3d.Architecture.Routing.Extensions ;
using Arent3d.Architecture.Routing.Storable ;
using Arent3d.Revit ;
using Autodesk.Revit.DB ;

namespace Arent3d.Architecture.Routing
{
  public static class ViewExtensions
  {
    private const string RoutingFamilyName = "Arent-Generic Models-Box" ;

    private const string RoutingViewPostFix = "Routing Assist" ;
    
    private const string View3DName = "3D";
    
    private const string Filter3DName = "3DModels";

    public static void CreateRoutingView( this Document document, IReadOnlyCollection<(ElementId Id, string Name)> levels )
    {
      var floorPlanFamily = document.GetAllElements<ViewFamilyType>().FirstOrDefault( viewFamilyType => viewFamilyType.ViewFamily == ViewFamily.FloorPlan ) ?? throw new InvalidOperationException() ;
      var views = document.GetAllElements<View>() ;
      var viewNames = views.Select( x => x.Name ).ToList() ;

      foreach ( var (id, name) in levels ) {
        var viewName = $"{name} - {RoutingViewPostFix}" ;
        if ( ! viewNames.Contains( viewName ) ) {
          var view = ViewPlan.Create( document, floorPlanFamily.Id, id ) ;
          foreach ( Category cat in view.Document.Settings.Categories ) {
            if ( cat.get_AllowsVisibilityControl( view ) ) {
              cat.set_Visible( view, IsViewable( cat ) ) ;
            }
          }

          view.Name = viewName ;
          view.ViewTemplateId = ElementId.InvalidElementId ;
          view.get_Parameter( BuiltInParameter.VIEW_DISCIPLINE ).Set( 4095 ) ;

          var pvr = view.GetViewRange() ;

          // pvr.SetLevelId(PlanViewPlane.TopClipPlane, vp.LevelId);
          pvr.SetOffset( PlanViewPlane.TopClipPlane, 4000.0 / 304.8 ) ;

          pvr.SetOffset( PlanViewPlane.CutPlane, 3000.0 / 304.8 ) ;

          // pvr.SetLevelId(PlanViewPlane.BottomClipPlane, vp.LevelId);
          pvr.SetOffset( PlanViewPlane.BottomClipPlane, 0.0 ) ;
          view.SetViewRange( pvr ) ;
        }
      }

      var filter = CreateElementFilter<RoutingFamilyType>( document, "RoutingModels" + DateTime.Now.ToString( "hhmmss" ) ) ;

      foreach ( View v in views ) {
        if ( NotContain( v.Name, RoutingViewPostFix ) ) {
          try {
            v.AddFilter( filter.Id ) ;
            v.SetFilterVisibility( filter.Id, false ) ;
          }
          catch {
            // ignored
          }
        }
      }
    }

    private static FilterElement CreateElementFilter<TFamilyTypeEnum>( Document document, string filterName ) where TFamilyTypeEnum : Enum
    {
      var familyNameParamId = new ElementId( BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM ) ;

      var categoryFilters = new List<ElementId>() ;
      var elementFilters = new List<ElementFilter>() ;

      foreach ( var field in typeof( TFamilyTypeEnum ).GetFields() ) {
        var nameOnRevit = field.GetCustomAttribute<NameOnRevitAttribute>() ;
        if ( null == nameOnRevit ) continue ;

        var familyCategory = field.GetCustomAttribute<FamilyCategoryAttribute>() ;
        if ( null == familyCategory ) continue ;

        var filterRule = ParameterFilterRuleFactory.CreateEqualsRule( familyNameParamId, RoutingFamilyName, true ) ;
        elementFilters.Add( new ElementParameterFilter( filterRule ) ) ;

        categoryFilters.Add( new ElementId( familyCategory.Category ) ) ;
      }


      var filter = ParameterFilterElement.Create( document, filterName, categoryFilters ) ;
      filter.SetElementFilter( new LogicalOrFilter( elementFilters ) ) ;

      return filter ;
    }

    private static bool NotContain( string s, string key )
    {
      if ( s.Contains( key ) ) {
        return false ;
      }

      return true ;
    }

    private static readonly BuiltInCategory[] ViewableElementCategory =
    {
      BuiltInCategory.OST_DuctAccessory,
      BuiltInCategory.OST_DuctCurves,
      BuiltInCategory.OST_DuctFitting,
      BuiltInCategory.OST_DuctTerminal,
      BuiltInCategory.OST_FlexDuctCurves,
      BuiltInCategory.OST_FlexPipeCurves,
      BuiltInCategory.OST_PipeCurves,
      BuiltInCategory.OST_PipeFitting,
      BuiltInCategory.OST_PipeSegments,
      BuiltInCategory.OST_GenericModel,
      BuiltInCategory.OST_MechanicalEquipment,
      BuiltInCategory.OST_PlumbingFixtures,
      BuiltInCategory.OST_Columns,
      BuiltInCategory.OST_Walls,
      BuiltInCategory.OST_Doors,
      BuiltInCategory.OST_Grids,
      BuiltInCategory.OST_Windows,
      BuiltInCategory.OST_StructuralFraming,
      BuiltInCategory.OST_Sections,
      BuiltInCategory.OST_SectionBox
    } ;

    private static bool IsViewable( Category cat )
    {
      return ViewableElementCategory.Contains( (BuiltInCategory) cat.Id.IntegerValue ) ;
    }

    public static void Create3DView( this Document document, IReadOnlyCollection<(ElementId Id, string Name)> levels )
    {
      if ( levels == null || levels.Count < 1 ) return ;

      HeightSettingStorable settingStorables = document.GetHeightSettingStorable() ;
      var allLevels = document.GetAllElements<Level>().OfCategory( BuiltInCategory.OST_Levels ).OrderBy( l => l.Elevation ).ToList() ;

      var viewFamilyType = document.GetAllElements<ViewFamilyType>().FirstOrDefault( viewFamilyType => viewFamilyType.ViewFamily == ViewFamily.ThreeDimensional ) ?? throw new InvalidOperationException() ;
      var views = document.GetAllElements<View>() ;
      var viewNames = views.Select( x => x.Name ).ToList() ;

      foreach ( var level in levels ) {
        var viewName = $"{View3DName}{level.Name}" ;
        if ( ! viewNames.Contains( viewName ) ) {
          var currentLevel = allLevels.FirstOrDefault( x => x.Id == level.Id ) ;
          var underFloor = settingStorables == null ? 0 : settingStorables[ currentLevel ].Underfloor.MillimetersToRevitUnits() ;
          View3D view = View3D.CreateIsometric( document, viewFamilyType.Id ) ;

          view.Name = viewName ;
          view.ViewTemplateId = ElementId.InvalidElementId ;
          view.DisplayStyle = DisplayStyle.FlatColors ;
          view.get_Parameter( BuiltInParameter.VIEW_DISCIPLINE ).Set( 4095 ) ;

          // Create a new BoundingBoxXYZ to define a 3D rectangular space
          BoundingBoxXYZ boundingBoxXYZ = new BoundingBoxXYZ() ;

          // Set the lower left bottom corner of the box
          // Use the Z of the current level.
          // X & Y values have been hardcoded based on this RVT geometry
          boundingBoxXYZ.Min = new XYZ( -50, -50, currentLevel.Elevation + underFloor ) ;

          // Determine the height of the bounding box
          double zOffset = 0 ;
          // If this is the top level, use an offset of 15 feet
          if ( level.Id == allLevels.Last().Id )
            zOffset = currentLevel.Elevation + 15 ;
          else {
            var aboveLevel = allLevels.Last() ;
            for ( int i = 0 ; i < allLevels.Count() - 1 ; i++ ) {
              if ( allLevels.ElementAt( i ).Id == level.Id ) {
                aboveLevel = allLevels.ElementAt( i + 1 ) ;
                break ;
              }
            }

            zOffset = aboveLevel.Elevation - 1 ;
          }

          boundingBoxXYZ.Max = new XYZ( 350, 250, zOffset ) ;

          // Apply this bounding box to the view's section box
          view.SetSectionBox( boundingBoxXYZ ) ;
        }
      }

      var filter = CreateElementFilter<RoutingFamilyType>( document, $"{Filter3DName}{DateTime.Now.ToString( "hhmmss" )}" ) ;

      foreach ( View v in views ) {
        if ( NotContain( v.Name, View3DName ) ) {
          try {
            v.AddFilter( filter.Id ) ;
            v.SetFilterVisibility( filter.Id, false ) ;
          }
          catch {
            // ignored
          }
        }
      }
    }
  }
}