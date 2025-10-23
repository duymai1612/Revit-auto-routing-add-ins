using System ;
using System.Collections.Generic ;
using System.Linq ;
using System.Windows ;
using Autodesk.Revit.DB ;
using System.Collections.ObjectModel ;
using System.Windows.Controls ;
using Arent3d.Revit ;
using Arent3d.Revit.I18n ;
using Arent3d.Utility ;
using ControlLib ;
using LengthConverter = Arent3d.Architecture.Routing.AppBase.Forms.ValueConverters.LengthConverter ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public partial class FromToEditControl : UserControl
  {
    private const string DefaultCurveTypeLabel = "Type" ;

    public event EventHandler? ValueChanged ;

    private void OnValueChanged( EventArgs e )
    {
      CanApply = CheckCanApply() ;
      IsChanged = CanApply && CheckIsChanged() ;
      ValueChanged?.Invoke( this, e ) ;
    }

    public static readonly DependencyProperty SystemTypeEditableProperty = DependencyProperty.Register( "SystemTypeEditable", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( true ) ) ;
    public static readonly DependencyProperty ShaftEditableProperty = DependencyProperty.Register( "ShaftEditable", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( true ) ) ;
    public static readonly DependencyProperty CurveTypeEditableProperty = DependencyProperty.Register( "CurveTypeEditable", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( true ) ) ;
    public static readonly DependencyProperty UseSystemTypeProperty = DependencyProperty.Register( "UseSystemType", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( true ) ) ;
    public static readonly DependencyProperty UseShaftProperty = DependencyProperty.Register( "UseShaft", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( true ) ) ;
    public static readonly DependencyProperty UseCurveTypeProperty = DependencyProperty.Register( "UseCurveType", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( true ) ) ;
    public static readonly DependencyProperty DiameterIndexProperty = DependencyProperty.Register( "DiameterIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( -1 ) ) ;
    public static readonly DependencyProperty SystemTypeIndexProperty = DependencyProperty.Register( "SystemTypeIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( -1 ) ) ;
    public static readonly DependencyProperty ShaftIndexProperty = DependencyProperty.Register( "ShaftIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( -1 ) ) ;
    public static readonly DependencyProperty CurveTypeIndexProperty = DependencyProperty.Register( "CurveTypeIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( -1 ) ) ;
    public static readonly DependencyProperty CurveTypeLabelProperty = DependencyProperty.Register( "CurveTypeLabel", typeof( string ), typeof( FromToEditControl ), new PropertyMetadata( DefaultCurveTypeLabel ) ) ;
    public static readonly DependencyProperty IsRouteOnPipeSpaceProperty = DependencyProperty.Register( "IsRouteOnPipeSpace", typeof( bool? ), typeof( FromToEditControl ), new PropertyMetadata( (bool?)true ) ) ;
    public static readonly DependencyProperty UseFromFixedHeightProperty = DependencyProperty.Register( "UseFromFixedHeight", typeof( bool? ), typeof( FromToEditControl ), new PropertyMetadata( (bool?)false ) ) ;
    public static readonly DependencyProperty FromFixedHeightProperty = DependencyProperty.Register( "FromFixedHeight", typeof( double? ), typeof( FromToEditControl ), new PropertyMetadata( 0.0, FromFixedHeight_Changed ) ) ;
    public static readonly DependencyProperty FromLocationTypeIndexProperty = DependencyProperty.Register( "FromLocationTypeIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( 0, FromLocationTypeIndex_PropertyChanged ) ) ;
    public static readonly DependencyProperty UseToFixedHeightProperty = DependencyProperty.Register( "UseToFixedHeight", typeof( bool? ), typeof( FromToEditControl ), new PropertyMetadata( (bool?)false ) ) ;
    public static readonly DependencyProperty ToFixedHeightProperty = DependencyProperty.Register( "ToFixedHeight", typeof( double? ), typeof( FromToEditControl ), new PropertyMetadata( 0.0, ToFixedHeight_Changed ) ) ;
    public static readonly DependencyProperty ToLocationTypeIndexProperty = DependencyProperty.Register( "ToLocationTypeIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( 0, ToLocationTypeIndex_PropertyChanged ) ) ;
    public static readonly DependencyProperty AvoidTypeIndexProperty = DependencyProperty.Register( "AvoidTypeIndex", typeof( int ), typeof( FromToEditControl ), new PropertyMetadata( 0 ) ) ;
    private static readonly DependencyPropertyKey CanApplyPropertyKey = DependencyProperty.RegisterReadOnly( "CanApply", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( false ) ) ;
    private static readonly DependencyPropertyKey IsChangedPropertyKey = DependencyProperty.RegisterReadOnly( "IsChanged", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( false ) ) ;
    private static readonly DependencyPropertyKey IsDifferentLevelPropertyKey = DependencyProperty.RegisterReadOnly( "IsDifferentLevel", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( false ) ) ;
    public static readonly DependencyProperty AllowIndeterminateProperty = DependencyProperty.Register( "AllowIndeterminate", typeof( bool ), typeof( FromToEditControl ), new PropertyMetadata( default( bool ) ) ) ;
    public static readonly DependencyProperty DisplayUnitSystemProperty = DependencyProperty.Register( "DisplayUnitSystem", typeof( DisplayUnit ), typeof( FromToEditControl ), new PropertyMetadata( DisplayUnit.IMPERIAL ) ) ;
    private static readonly DependencyPropertyKey FromMinimumHeightAsFloorLevelPropertyKey = DependencyProperty.RegisterReadOnly( "FromMinimumHeightAsFloorLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey FromMaximumHeightAsFloorLevelPropertyKey = DependencyProperty.RegisterReadOnly( "FromMaximumHeightAsFloorLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey FromMinimumHeightAsCeilingLevelPropertyKey = DependencyProperty.RegisterReadOnly( "FromMinimumHeightAsCeilingLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey FromMaximumHeightAsCeilingLevelPropertyKey = DependencyProperty.RegisterReadOnly( "FromMaximumHeightAsCeilingLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey FromDefaultHeightAsFloorLevelPropertyKey = DependencyProperty.RegisterReadOnly( "FromDefaultHeightAsFloorLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey FromDefaultHeightAsCeilingLevelPropertyKey = DependencyProperty.RegisterReadOnly( "FromDefaultHeightAsCeilingLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey ToMinimumHeightAsFloorLevelPropertyKey = DependencyProperty.RegisterReadOnly( "ToMinimumHeightAsFloorLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey ToMaximumHeightAsFloorLevelPropertyKey = DependencyProperty.RegisterReadOnly( "ToMaximumHeightAsFloorLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey ToMinimumHeightAsCeilingLevelPropertyKey = DependencyProperty.RegisterReadOnly( "ToMinimumHeightAsCeilingLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey ToMaximumHeightAsCeilingLevelPropertyKey = DependencyProperty.RegisterReadOnly( "ToMaximumHeightAsCeilingLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey ToDefaultHeightAsFloorLevelPropertyKey = DependencyProperty.RegisterReadOnly( "ToDefaultHeightAsFloorLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;
    private static readonly DependencyPropertyKey ToDefaultHeightAsCeilingLevelPropertyKey = DependencyProperty.RegisterReadOnly( "ToDefaultHeightAsCeilingLevel", typeof( double ), typeof( FromToEditControl ), new PropertyMetadata( 0.0 ) ) ;

    //Diameter Info
    private double VertexTolerance { get ; set ; }
    public ObservableCollection<double> Diameters { get ; } = new ObservableCollection<double>() ;
    private double? DiameterOrg { get ; set ; }

    public double? Diameter
    {
      get => GetDiameterOnIndex( Diameters, (int)GetValue( DiameterIndexProperty ) ) ;
      private set => SetValue( DiameterIndexProperty, GetDiameterIndex( Diameters, value, VertexTolerance ) ) ;
    }

    private static double? GetDiameterOnIndex( IReadOnlyList<double> diameters, int index )
    {
      if ( index < 0 || diameters.Count <= index ) return null ;
      return diameters[ index ] ;
    }

    private static int GetDiameterIndex( IReadOnlyList<double> diameters, double? value, double tolerance )
    {
      if ( value is not { } diameter ) {
        if ( 0 < diameters.Count ) return 0 ; // Use minimum value
        return -1 ;
      }

      return diameters.FindIndex( d => LengthEquals( d, diameter, tolerance ) ) ;
    }

    private static bool LengthEquals( double d1, double d2, double tolerance )
    {
      return Math.Abs( d1 - d2 ) < tolerance ;
    }

    private static bool LengthEquals( double? d1, double? d2, double tolerance )
    {
      if ( d1.HasValue != d2.HasValue ) return false ;
      if ( false == d1.HasValue ) return true ;

      return LengthEquals( d1.Value, d2!.Value, tolerance ) ;
    }

    //SystemType Info
    public ObservableCollection<MEPSystemType> SystemTypes { get ; } = new ObservableCollection<MEPSystemType>() ;
    private MEPSystemType? SystemTypeOrg { get ; set ; }

    public MEPSystemType? SystemType
    {
      get => GetItemOnIndex( SystemTypes, (int)GetValue( SystemTypeIndexProperty ) ) ;
      private set => SetValue( SystemTypeIndexProperty, GetItemIndex( SystemTypes, value ) ) ;
    }

    public bool SystemTypeEditable
    {
      get => (bool)GetValue( SystemTypeEditableProperty ) ;
      set => SetValue( SystemTypeEditableProperty, value ) ;
    }

    private bool UseSystemType
    {
      get => (bool)GetValue( UseSystemTypeProperty ) ;
      set => SetValue( UseSystemTypeProperty, value ) ;
    }

    //Shafts Info
    public ObservableCollection<OpeningProxy> Shafts { get ; } = new ObservableCollection<OpeningProxy>() ;
    private Opening? ShaftOrg { get ; set ; }

    public Opening? Shaft
    {
      get => GetItemOnIndex( Shafts, (int)GetValue( ShaftIndexProperty ) )?.Value ;
      private set => SetValue( ShaftIndexProperty, GetShaftIndex( Shafts, value ) ) ;
    }

    public bool ShaftEditable
    {
      get => (bool)GetValue( ShaftEditableProperty ) ;
      set => SetValue( ShaftEditableProperty, value ) ;
    }

    private bool UseShaft
    {
      get => (bool)GetValue( UseShaftProperty ) ;
      set => SetValue( UseShaftProperty, value ) ;
    }

    //CurveType Info
    public ObservableCollection<MEPCurveType> CurveTypes { get ; } = new ObservableCollection<MEPCurveType>() ;
    private MEPCurveType? CurveTypeOrg { get ; set ; }

    public MEPCurveType? CurveType
    {
      get => GetItemOnIndex( CurveTypes, (int)GetValue( CurveTypeIndexProperty ) ) ;
      private set
      {
        SetValue( CurveTypeIndexProperty, GetItemIndex( CurveTypes, value ) ) ;
        if ( value is { } curveType ) {
          CurveTypeLabel = UIHelper.GetTypeLabel( curveType.GetType().Name ) ;
        }
        else {
          CurveTypeLabel = DefaultCurveTypeLabel ;
        }

        UpdateDiameterList() ;
      }
    }

    private void UpdateDiameterList()
    {
      var curveType = CurveType ;
      var currentDiameter = Diameter ;

      Diameters.Clear() ;
      if ( curveType?.GetNominalDiameters( VertexTolerance ) is { } diameters ) {
        diameters.ForEach( Diameters.Add ) ;
      }

      if ( currentDiameter is { } d ) {
        SetCurrentValue( DiameterIndexProperty, UIHelper.FindClosestIndex( Diameters, d ) ) ;
      }
      else {
        SetCurrentValue( DiameterIndexProperty, -1 ) ;
      }
    }

    public DisplayUnit DisplayUnitSystem
    {
      get { return (DisplayUnit)GetValue( DisplayUnitSystemProperty ) ; }
      set { SetValue( DisplayUnitSystemProperty, value ) ; }
    }

    private string CurveTypeLabel
    {
      get => (string)GetValue( CurveTypeLabelProperty ) ;
      set => SetValue( CurveTypeLabelProperty, value ) ;
    }

    public bool CurveTypeEditable
    {
      get => (bool)GetValue( CurveTypeEditableProperty ) ;
      set => SetValue( CurveTypeEditableProperty, value ) ;
    }

    private bool UseCurveType
    {
      get => (bool)GetValue( UseCurveTypeProperty ) ;
      set => SetValue( UseCurveTypeProperty, value ) ;
    }

    private static T? GetItemOnIndex<T>( IReadOnlyList<T> values, int index ) where T : class
    {
      if ( index < 0 || values.Count <= index ) return null ;
      return values[ index ] ;
    }

    private static int GetItemIndex<TElement>( IEnumerable<TElement> elements, TElement? value ) where TElement : Element
    {
      var valueId = value.GetValidId() ;
      if ( ElementId.InvalidElementId == valueId ) return -1 ;

      return elements.FindIndex( elm => elm.Id == valueId ) ;
    }

    private static int GetShaftIndex( IEnumerable<OpeningProxy> elements, Opening? value )
    {
      var valueId = value.GetValidId() ;
      return elements.FindIndex( elm => elm.Value?.Id == valueId ) ;
    }

    //Direct Info
    private bool? IsRouteOnPipeSpaceOrg { get ; set ; }

    public bool? IsRouteOnPipeSpace
    {
      get => (bool?)GetValue( IsRouteOnPipeSpaceProperty ) ;
      private set => SetValue( IsRouteOnPipeSpaceProperty, value ) ;
    }

    //HeightSetting
    private bool? UseFromFixedHeightOrg { get ; set ; }
    private double? FromFixedHeightOrg { get ; set ; }

    public bool? UseFromFixedHeight
    {
      get => (bool?)GetValue( UseFromFixedHeightProperty ) ;
      private set => SetValue( UseFromFixedHeightProperty, value ) ;
    }

    public double? FromFixedHeight
    {
      get => (double?)GetValue( FromFixedHeightProperty ) ;
      private set => SetValue( FromFixedHeightProperty, value ) ;
    }

    //ToHeightSetting
    private bool? UseToFixedHeightOrg { get ; set ; }
    private double? ToFixedHeightOrg { get ; set ; }

    public bool? UseToFixedHeight
    {
      get => (bool?)GetValue( UseToFixedHeightProperty ) ;
      private set => SetValue( UseToFixedHeightProperty, value ) ;
    }

    public double? ToFixedHeight
    {
      get => (double?)GetValue( ToFixedHeightProperty ) ;
      private set => SetValue( ToFixedHeightProperty, value ) ;
    }

    private double FromMinimumHeightAsFloorLevel
    {
      get => (double)GetValue( FromMinimumHeightAsFloorLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( FromMinimumHeightAsFloorLevelPropertyKey, value ) ;
    }
    private double FromMaximumHeightAsFloorLevel
    {
      get => (double)GetValue( FromMaximumHeightAsFloorLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( FromMaximumHeightAsFloorLevelPropertyKey, value ) ;
    }
    private double FromMinimumHeightAsCeilingLevel
    {
      get => (double)GetValue( FromMinimumHeightAsCeilingLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( FromMinimumHeightAsCeilingLevelPropertyKey, value ) ;
    }
    private double FromMaximumHeightAsCeilingLevel
    {
      get => (double)GetValue( FromMaximumHeightAsCeilingLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( FromMaximumHeightAsCeilingLevelPropertyKey, value ) ;
    }
    private double FromDefaultHeightAsFloorLevel
    {
      get => (double)GetValue( FromDefaultHeightAsFloorLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( FromDefaultHeightAsFloorLevelPropertyKey, value ) ;
    }
    private double FromDefaultHeightAsCeilingLevel
    {
      get => (double)GetValue( FromDefaultHeightAsCeilingLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( FromDefaultHeightAsCeilingLevelPropertyKey, value ) ;
    }
    private double ToMinimumHeightAsFloorLevel
    {
      get => (double)GetValue( ToMinimumHeightAsFloorLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( ToMinimumHeightAsFloorLevelPropertyKey, value ) ;
    }
    private double ToMaximumHeightAsFloorLevel
    {
      get => (double)GetValue( ToMaximumHeightAsFloorLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( ToMaximumHeightAsFloorLevelPropertyKey, value ) ;
    }
    private double ToMinimumHeightAsCeilingLevel
    {
      get => (double)GetValue( ToMinimumHeightAsCeilingLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( ToMinimumHeightAsCeilingLevelPropertyKey, value ) ;
    }
    private double ToMaximumHeightAsCeilingLevel
    {
      get => (double)GetValue( ToMaximumHeightAsCeilingLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( ToMaximumHeightAsCeilingLevelPropertyKey, value ) ;
    }
    private double ToDefaultHeightAsFloorLevel
    {
      get => (double)GetValue( ToDefaultHeightAsFloorLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( ToDefaultHeightAsFloorLevelPropertyKey, value ) ;
    }
    private double ToDefaultHeightAsCeilingLevel
    {
      get => (double)GetValue( ToDefaultHeightAsCeilingLevelPropertyKey.DependencyProperty ) ;
      set => SetValue( ToDefaultHeightAsCeilingLevelPropertyKey, value ) ;
    }

    private static void FromLocationTypeIndex_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as FromToEditControl )?.OnFromLocationTypeChanged() ;
    }
    private static void ToLocationTypeIndex_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as FromToEditControl )?.OnToLocationTypeChanged() ;
    }
    private void OnFromLocationTypeChanged()
    {
      if ( FromLocationType is not { } locationType ) return ;

      var minimumValue = ( locationType == FixedHeightType.Ceiling ? FromMinimumHeightAsCeilingLevel : FromMinimumHeightAsFloorLevel ) ;
      var maximumValue = ( locationType == FixedHeightType.Ceiling ? FromMaximumHeightAsCeilingLevel : FromMaximumHeightAsFloorLevel ) ;
      SetMinMax( FromFixedHeightNumericUpDown, locationType, minimumValue, maximumValue ) ;
    }
    private void OnToLocationTypeChanged()
    {
      if ( ToLocationType is not { } locationType ) return ;

      var minimumValue = ( locationType == FixedHeightType.Ceiling ? ToMinimumHeightAsCeilingLevel : ToMinimumHeightAsFloorLevel ) ;
      var maximumValue = ( locationType == FixedHeightType.Ceiling ? ToMaximumHeightAsCeilingLevel : ToMaximumHeightAsFloorLevel ) ;
      SetMinMax( ToFixedHeightNumericUpDown, locationType, minimumValue, maximumValue ) ;
    }

    private void SetMinMax( NumericUpDown numericUpDown, FixedHeightType locationType, double minimumValue, double maximumValue )
    {
      var lengthConverter = GetLengthConverter( DisplayUnitSystem ) ;
      numericUpDown.MinValue = Math.Round( lengthConverter.ConvertUnit( minimumValue ), 5, MidpointRounding.AwayFromZero ) ;
      numericUpDown.MaxValue = Math.Round( lengthConverter.ConvertUnit( maximumValue ), 5, MidpointRounding.AwayFromZero ) ;
      numericUpDown.Value = Math.Max( numericUpDown.MinValue, Math.Min( numericUpDown.Value, numericUpDown.MaxValue ) ) ;
      numericUpDown.ToolTip = $"{numericUpDown.MinValue} ～ {numericUpDown.MaxValue}" ;
    }

    //AvoidType
    private AvoidType? AvoidTypeOrg { get ; set ; }

    public AvoidType? AvoidType
    {
      get => GetAvoidTypeOnIndex( AvoidTypes.Keys, (int)GetValue( AvoidTypeIndexProperty ) ) ;
      private set => SetValue( AvoidTypeIndexProperty, GetAvoidTypeIndex( AvoidTypes.Keys, value ) ) ;
    }

    private static AvoidType? GetAvoidTypeOnIndex( IEnumerable<AvoidType> avoidTypes, int index )
    {
      if ( index < 0 ) return null ;
      return avoidTypes.ElementAtOrDefault( index ) ;
    }

    private static int GetAvoidTypeIndex( IEnumerable<AvoidType> avoidTypes, AvoidType? avoidType )
    {
      return ( avoidType is { } type ? avoidTypes.IndexOf( type ) : -1 ) ;
    }

    public IReadOnlyDictionary<AvoidType, string> AvoidTypes { get ; } = new Dictionary<AvoidType, string>
    {
      [ Routing.AvoidType.Whichever ] = "Dialog.Forms.FromToEditControl.ProcessConstraints.None".GetAppStringByKeyOrDefault( "Whichever" ),
      [ Routing.AvoidType.NoAvoid ] = "Dialog.Forms.FromToEditControl.ProcessConstraints.NoPocket".GetAppStringByKeyOrDefault( "Don't avoid From-To" ),
      [ Routing.AvoidType.AvoidAbove ] = "Dialog.Forms.FromToEditControl.ProcessConstraints.NoDrainPocket".GetAppStringByKeyOrDefault( "Avoid on From-To" ),
      [ Routing.AvoidType.AvoidBelow ] = "Dialog.Forms.FromToEditControl.ProcessConstraints.NoVentPocket".GetAppStringByKeyOrDefault( "Avoid below From-To" ),
    } ;

    //LocationType
    private FixedHeightType? FromLocationTypeOrg { get ; set ; }

    public FixedHeightType? FromLocationType
    {
      get => GetLocationTypeOnIndex( LocationTypes.Keys, (int)GetValue( FromLocationTypeIndexProperty ) ) ;
      private set => SetValue( FromLocationTypeIndexProperty, GetLocationTypeIndex( LocationTypes.Keys, value ) ) ;
    }

    private FixedHeightType? ToLocationTypeOrg { get ; set ; }

    public FixedHeightType? ToLocationType
    {
      get => GetLocationTypeOnIndex( LocationTypes.Keys, (int)GetValue( ToLocationTypeIndexProperty ) ) ;
      private set => SetValue( ToLocationTypeIndexProperty, GetLocationTypeIndex( LocationTypes.Keys, value ) ) ;
    }

    private static FixedHeightType? GetLocationTypeOnIndex( IEnumerable<FixedHeightType> locationTypes, int index )
    {
      if ( index < 0 ) return null ;
      return locationTypes.ElementAtOrDefault( index ) ;
    }

    private static int GetLocationTypeIndex( IEnumerable<FixedHeightType> locationTypes, FixedHeightType? locationType )
    {
      return ( locationType is { } type ? locationTypes.IndexOf( type ) : -1 ) ;
    }

    public IReadOnlyDictionary<FixedHeightType, string> LocationTypes { get ; } = new Dictionary<FixedHeightType, string>
    {
      [ FixedHeightType.Floor ] = "FL",
      [ FixedHeightType.Ceiling] = "CL",
    } ;

    public bool IsDifferentLevel
    {
      get => (bool)GetValue( IsDifferentLevelPropertyKey.DependencyProperty ) ;
      private set => SetValue( IsDifferentLevelPropertyKey, value ) ;
    }

    public bool CanApply
    {
      get => (bool)GetValue( CanApplyPropertyKey.DependencyProperty ) ;
      private set => SetValue( CanApplyPropertyKey, value ) ;
    }

    public bool IsChanged
    {
      get => (bool)GetValue( IsChangedPropertyKey.DependencyProperty ) ;
      private set => SetValue( IsChangedPropertyKey, value ) ;
    }

    public bool AllowIndeterminate
    {
      get { return (bool)GetValue( AllowIndeterminateProperty ) ; }
      set { SetValue( AllowIndeterminateProperty, value ) ; }
    }

    private bool CheckCanApply()
    {
      if ( false == AllowIndeterminate ) {
        if ( UseSystemType && null == SystemType ) return false ;
        if ( UseCurveType && null == CurveType ) return false ;
        if ( null == Diameter ) return false ;
        if ( null == IsRouteOnPipeSpace ) return false ;
        if ( null == UseFromFixedHeight ) return false ;
        if ( null == FromLocationType ) return false ;
        if ( null == FromFixedHeight ) return false ;
        if ( IsDifferentLevel ) {
          if ( null == UseToFixedHeight ) return false ;
          if ( null == ToLocationType ) return false ;
          if ( null == ToFixedHeight ) return false ;
        }
      }

      return true ;
    }

    private bool CheckIsChanged()
    {
      if ( UseSystemType && SystemTypeOrg.GetValidId() != SystemType.GetValidId() ) return true ;
      if ( UseCurveType && CurveTypeOrg.GetValidId() != CurveType.GetValidId() ) return true ;
      if ( false == LengthEquals( DiameterOrg, Diameter, VertexTolerance ) ) return true ;
      if ( IsRouteOnPipeSpace != IsRouteOnPipeSpaceOrg ) return true ;
      if ( UseFromFixedHeight != UseFromFixedHeightOrg ) return true ;
      if ( true == UseFromFixedHeight ) {
        if ( FromLocationTypeOrg != FromLocationType ) return true ;
        if ( false == LengthEquals( FromFixedHeightOrg, FromFixedHeight, VertexTolerance ) ) return true ;
      }
      if ( IsDifferentLevel ) {
        if ( true == UseToFixedHeight ) {
          if ( ToLocationTypeOrg != ToLocationType ) return true ;
          if ( false == LengthEquals( ToFixedHeightOrg, ToFixedHeight, VertexTolerance ) ) return true ;
        }
      }
      if ( AvoidTypeOrg != AvoidType ) return true ;
      if ( UseShaft && ShaftOrg.GetValidId() != Shaft.GetValidId() ) return true ;

      return false ;
    }

    public FromToEditControl()
    {
      InitializeComponent() ;

      ClearDialog() ;
    }

    private void SystemTypeComboBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void ShaftComboBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void CurveTypeComboBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }


    private void DiameterComboBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void SetAvailableParameterList( RoutePropertyTypeList propertyTypeList )
    {
      Diameters.Clear() ;
      SystemTypes.Clear() ;
      CurveTypes.Clear() ;
      Shafts.Clear() ;

      // System type
      if ( propertyTypeList.SystemTypes is {} systemTypes ) {
        foreach ( var s in systemTypes ) {
          SystemTypes.Add( s ) ;
        }

        UseSystemType = true ;
      }
      else {
        UseSystemType = false ;
      }

      Shafts.Add( new OpeningProxy( null ) ) ;
      if ( propertyTypeList.Shafts is {} shafts ) {
        foreach ( var shaft in shafts ) {
          Shafts.Add( new OpeningProxy( shaft ) ) ;
        }

        UseShaft = true ;
      }
      else {
        UseShaft = false ;
      }

      // Curve type
      foreach ( var c in propertyTypeList.CurveTypes ) {
        CurveTypes.Add( c ) ;
      }

      IsDifferentLevel = propertyTypeList.HasDifferentLevel ;

      ( FromMinimumHeightAsFloorLevel, FromMaximumHeightAsFloorLevel ) = propertyTypeList.FromHeightRangeAsFloorLevel ;
      ( FromMinimumHeightAsCeilingLevel, FromMaximumHeightAsCeilingLevel ) = propertyTypeList.FromHeightRangeAsCeilingLevel ;
      FromDefaultHeightAsFloorLevel = propertyTypeList.FromDefaultHeightAsFloorLevel ;
      FromDefaultHeightAsCeilingLevel = propertyTypeList.FromDefaultHeightAsCeilingLevel ;

      ( ToMinimumHeightAsFloorLevel, ToMaximumHeightAsFloorLevel ) = propertyTypeList.ToHeightRangeAsFloorLevel ;
      ( ToMinimumHeightAsCeilingLevel, ToMaximumHeightAsCeilingLevel ) = propertyTypeList.ToHeightRangeAsCeilingLevel ;
      ToDefaultHeightAsFloorLevel = propertyTypeList.ToDefaultHeightAsFloorLevel ;
      ToDefaultHeightAsCeilingLevel = propertyTypeList.ToDefaultHeightAsCeilingLevel ;
    }

    public void SetRouteProperties( RoutePropertyTypeList propertyTypeList, RouteProperties properties )
    {
      VertexTolerance = properties.VertexTolerance ;
      SetAvailableParameterList( propertyTypeList ) ;

      SystemTypeOrg = properties.SystemType ;
      CurveTypeOrg = properties.CurveType ;
      DiameterOrg = properties.Diameter ;
      ShaftOrg = properties.Shaft ;

      IsRouteOnPipeSpaceOrg = properties.IsRouteOnPipeSpace ;

      UseFromFixedHeightOrg = properties.UseFromFixedHeight ;
      if ( null == UseFromFixedHeightOrg ) {
        FromLocationTypeOrg = null ;
        FromFixedHeightOrg = null ;
      }
      else {
        FromLocationTypeOrg = properties.FromFixedHeight?.Type ?? FixedHeightType.Ceiling ;
        FromFixedHeightOrg = properties.FromFixedHeight?.Height ?? GetFromDefaultHeight( FromLocationTypeOrg.Value ) ;
      }

      UseToFixedHeightOrg = properties.UseToFixedHeight ;
      if ( null == UseToFixedHeightOrg ) {
        ToLocationTypeOrg = null ;
        ToFixedHeightOrg = null ;
      }
      else {
        ToLocationTypeOrg = properties.ToFixedHeight?.Type ?? FixedHeightType.Ceiling ;
        ToFixedHeightOrg = properties.ToFixedHeight?.Height ?? GetFromDefaultHeight( ToLocationTypeOrg.Value ) ;
      }

      AvoidTypeOrg = properties.AvoidType ;
    }

    public void ResetDialog()
    {
      SystemType = SystemTypeOrg ;
      CurveType = CurveTypeOrg ;
      Diameter = DiameterOrg ;
      Shaft = ShaftOrg ;

      IsRouteOnPipeSpace = IsRouteOnPipeSpaceOrg ;

      UseFromFixedHeight = UseFromFixedHeightOrg ;
      FromLocationType = FromLocationTypeOrg ;
      FromFixedHeight = FromFixedHeightOrg ;
      UseToFixedHeight = UseToFixedHeightOrg ;
      ToLocationType = ToLocationTypeOrg ;
      ToFixedHeight = ToFixedHeightOrg ;

      OnFromLocationTypeChanged() ;
      OnToLocationTypeChanged() ;

      AvoidType = AvoidTypeOrg ;

      CanApply = false ;
    }

    public void ClearDialog()
    {
      Diameters.Clear() ;
      SystemTypes.Clear() ;
      CurveTypes.Clear() ;

      DiameterOrg = null ;
      SystemTypeOrg = null ;
      CurveTypeOrg = null ;
      ShaftOrg = null ;

      IsRouteOnPipeSpaceOrg = false ;

      UseFromFixedHeightOrg = false ;
      FromLocationTypeOrg = FixedHeightType.Floor ;
      FromFixedHeight = null ;
      UseToFixedHeightOrg = false ;
      ToLocationTypeOrg = FixedHeightType.Floor ;
      ToFixedHeight = null ;

      AvoidTypeOrg = Routing.AvoidType.Whichever ;

      ResetDialog() ;
    }

    private void Direct_OnChecked( object sender, RoutedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void Direct_OnUnchecked( object sender, RoutedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void Height_OnChecked( object sender, RoutedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void Height_OnUnchecked( object sender, RoutedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void ToHeight_OnChecked( object sender, RoutedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void ToHeight_OnUnchecked( object sender, RoutedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void AvoidTypeComboBox_OnSelectionChanged( object sender, SelectionChangedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;
    }

    private void LocationTypeComboBox_SelectionChanged( object sender, SelectionChangedEventArgs e )
    {
      OnValueChanged( EventArgs.Empty ) ;

      if ( e.RemovedItems.OfType<KeyValuePair<FixedHeightType, string>?>().FirstOrDefault() is not { } oldValue || e.AddedItems.OfType<KeyValuePair<FixedHeightType, string>?>().FirstOrDefault() is not { } newValue ) return ;
      if ( oldValue.Key == newValue.Key ) return ;

      if ( object.ReferenceEquals( sender, FromLocationTypeComboBox ) ) {
        FromFixedHeight = GetFromDefaultHeight( newValue.Key ) ;
      }
      else {
        ToFixedHeight = GetToDefaultHeight( newValue.Key ) ;
      }
    }

    private double? GetFromDefaultHeight( FixedHeightType newValue )
    {
      return ( FixedHeightType.Ceiling == newValue ) ? FromDefaultHeightAsCeilingLevel : FromDefaultHeightAsFloorLevel ;
    }
    private double? GetToDefaultHeight( FixedHeightType newValue )
    {
      return ( FixedHeightType.Ceiling == newValue ) ? ToDefaultHeightAsCeilingLevel : ToDefaultHeightAsFloorLevel ;
    }

    private void FromFixedHeightNumericUpDown_OnValueChanged( object sender, ValueChangedEventArgs e )
    {
      // Manually update FixedHeight because Value binding is not called.
      FromFixedHeight = GetLengthConverter( DisplayUnitSystem ).ConvertBackUnit( FromFixedHeightNumericUpDown.Value ) ;

      OnValueChanged( EventArgs.Empty ) ;
    }

    private void ToFixedHeightNumericUpDown_OnValueChanged( object sender, ValueChangedEventArgs e )
    {
      // Manually update FixedHeight because Value binding is not called.
      ToFixedHeight = GetLengthConverter( DisplayUnitSystem ).ConvertBackUnit( ToFixedHeightNumericUpDown.Value ) ;

      OnValueChanged( EventArgs.Empty ) ;
    }

    private static void FromFixedHeight_Changed( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      if ( d is not FromToEditControl fromToEditControl ) return ;

      if ( e.NewValue is not double newValue ) {
        fromToEditControl.FromFixedHeightNumericUpDown.CanHaveNull = true ;
        fromToEditControl.FromFixedHeightNumericUpDown.HasValidValue = false ;
      }
      else {
        fromToEditControl.FromFixedHeightNumericUpDown.CanHaveNull = false ;
        fromToEditControl.FromFixedHeightNumericUpDown.Value = Math.Round( GetLengthConverter( fromToEditControl.DisplayUnitSystem ).ConvertUnit( newValue ), 5, MidpointRounding.AwayFromZero ) ;
      }
    }

    private static void ToFixedHeight_Changed( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      if ( d is not FromToEditControl fromToEditControl ) return ;

      if ( e.NewValue is not double newValue ) {
        fromToEditControl.ToFixedHeightNumericUpDown.CanHaveNull = true ;
        fromToEditControl.ToFixedHeightNumericUpDown.HasValidValue = false ;
      }
      else {
        fromToEditControl.ToFixedHeightNumericUpDown.CanHaveNull = false ;
        fromToEditControl.ToFixedHeightNumericUpDown.Value = Math.Round( GetLengthConverter( fromToEditControl.DisplayUnitSystem ).ConvertUnit( newValue ), 5, MidpointRounding.AwayFromZero ) ;
      }
    }

    private static LengthConverter GetLengthConverter( DisplayUnit displayUnitSystem )
    {
      return displayUnitSystem switch
      {
        DisplayUnit.METRIC => LengthConverter.Millimeters,
        DisplayUnit.IMPERIAL => LengthConverter.Inches,
        _ => LengthConverter.Default,
      } ;
    }

    public class OpeningProxy
    {
      internal OpeningProxy( Opening? opening )
      {
        Value = opening ;
      }

      public Opening? Value { get ; }
    }
  }
}