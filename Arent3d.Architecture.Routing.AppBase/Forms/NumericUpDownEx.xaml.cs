using System ;
using System.Threading.Tasks ;
using System.Windows ;
using System.Windows.Controls ;
using System.Windows.Controls.Primitives ;
using System.Windows.Data ;
using System.Windows.Input ;
using ControlLib ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public partial class NumericUpDownEx : NumericUpDown
  {
    #region Dependency properties

    public static readonly DependencyProperty HasValidValueProperty = DependencyProperty.Register( nameof( HasValidValue ), typeof( bool ), typeof( NumericUpDownEx ), new PropertyMetadata( true, HasValidValue_PropertyChanged ) ) ;
    public static readonly DependencyProperty CanHaveNullProperty = DependencyProperty.Register( nameof( CanHaveNull ), typeof( bool ), typeof( NumericUpDownEx ), new PropertyMetadata( false, CanHaveNull_PropertyChanged ) ) ;
    public static readonly DependencyProperty TextAlignmentProperty = DependencyProperty.Register( nameof( TextAlignment ), typeof( TextAlignment ), typeof( NumericUpDownEx ), new PropertyMetadata( TextAlignment.Left ) ) ;

    public TextAlignment TextAlignment
    {
      get { return (TextAlignment)GetValue( TextAlignmentProperty ) ; }
      set { SetValue( TextAlignmentProperty, value ) ; }
    }

    public bool HasValidValue
    {
      get { return (bool)GetValue( HasValidValueProperty ) ; }
      set { SetValue( HasValidValueProperty, value ) ; }
    }

    public bool CanHaveNull
    {
      get { return (bool)GetValue( CanHaveNullProperty ) ; }
      set { SetValue( CanHaveNullProperty, value ) ; }
    }

    private static void HasValidValue_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as NumericUpDownEx )?.OnHasValidValueChanged( (bool)e.NewValue ) ;
    }

    private static void CanHaveNull_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as NumericUpDownEx )?.OnCanHaveNullChanged( (bool)e.NewValue ) ;
    }

    #endregion
    
    #region Overwriting dependency properties

    static NumericUpDownEx()
    {
      ValueProperty.OverrideMetadata( typeof( NumericUpDownEx ), new FrameworkPropertyMetadata( 0.0, ValueProperty_PropertyChanged, ValueProperty_CoerceValue ) ) ;
      MinValueProperty.OverrideMetadata( typeof( NumericUpDownEx ), new FrameworkPropertyMetadata( 0.0, MinValueProperty_PropertyChanged ) ) ;
      MaxValueProperty.OverrideMetadata( typeof( NumericUpDownEx ), new FrameworkPropertyMetadata( 100.0, MaxValueProperty_PropertyChanged ) ) ;
    }

    private static void ValueProperty_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as NumericUpDownEx )?.OnValueChanged( e ) ;
    }
    private static object ValueProperty_CoerceValue( DependencyObject d, object value )
    {
      return ( d as NumericUpDownEx )?.CoerceValueProperty( (double)value ) ?? value ;
    }

    private static void MinValueProperty_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as NumericUpDownEx )?.OnMinValueChanged( e ) ;
    }

    private static void MaxValueProperty_PropertyChanged( DependencyObject d, DependencyPropertyChangedEventArgs e )
    {
      ( d as NumericUpDownEx )?.OnMaxValueChanged( e ) ;
    }

    #endregion

    #region Routed events

    public static readonly RoutedEvent MinValueChangedEvent = EventManager.RegisterRoutedEvent( nameof( MinValueChanged ), RoutingStrategy.Direct, typeof( ValueChangedEventHandler ), typeof( NumericUpDownEx ) ) ;
    public static readonly RoutedEvent MaxValueChangedEvent = EventManager.RegisterRoutedEvent( nameof( MaxValueChanged ), RoutingStrategy.Direct, typeof( ValueChangedEventHandler ), typeof( NumericUpDownEx ) ) ;
    public static readonly RoutedEvent HasValidValueChangedEvent = EventManager.RegisterRoutedEvent( nameof( HasValidValueChanged ), RoutingStrategy.Direct, typeof( RoutedEventHandler ), typeof( NumericUpDownEx ) ) ;

    public event RoutedEventHandler HasValidValueChanged
    {
      add => AddHandler( HasValidValueChangedEvent, value ) ;
      remove => RemoveHandler( HasValidValueChangedEvent, value ) ;
    }

    public event ValueChangedEventHandler MinValueChanged
    {
      add => AddHandler( MinValueChangedEvent, value ) ;
      remove => RemoveHandler( MinValueChangedEvent, value ) ;
    }

    public event ValueChangedEventHandler MaxValueChanged
    {
      add => AddHandler( MaxValueChangedEvent, value ) ;
      remove => RemoveHandler( MaxValueChangedEvent, value ) ;
    }

    #endregion

    private bool _textToValue = false ;
    private TextBox? _textBox ;
    private ButtonBase? _up ;
    private ButtonBase? _down ;

    public NumericUpDownEx()
    {
      InitializeComponent() ;
    }

    public override void OnApplyTemplate()
    {
      if ( this.GetTemplateChild( "PART_TextBox" ) is TextBox textBox ) {
        _textBox = textBox ;
        _textBox.GotFocus += ( _, _ ) => SelectAll() ;
        _textBox.PreviewKeyDown += TextBox_PreviewKeyDown ;
        _textBox.TextChanged += TextBox_TextChanged ;
        _textBox.SetBinding( TextBox.TextAlignmentProperty, new Binding( nameof( TextAlignment ) ) { Source = this } ) ;
      }

      if ( this.GetTemplateChild( "PART_ButtonUp" ) is ButtonBase upButton ) {
        _up = upButton ;
        _up.Focusable = false ;
        _up.Click += ( _, _ ) => SelectAll() ;
      }

      if ( this.GetTemplateChild( "PART_ButtonDown" ) is ButtonBase downButton ) {
        _down = downButton ;
        _down.Focusable = false ;
        _down.Click += ( _, _ ) => SelectAll() ;
      }

      base.OnApplyTemplate() ;
    }

    private void OnHasValidValueChanged( bool newHasValidValue )
    {
      if ( false == CanHaveNull && false == newHasValidValue ) HasValidValue = true ;

      if ( null != _textBox && false == _textToValue ) {
        _textBox.Text = ( newHasValidValue ? Value.ToString() : string.Empty ) ;
      }

      RaiseEvent( new RoutedEventArgs( HasValidValueChangedEvent ) ) ;
    }

    private void OnCanHaveNullChanged( bool newCanHaveNull )
    {
      if ( false == newCanHaveNull ) {
        HasValidValue = true ;
      }
    }

    private void OnValueChanged( DependencyPropertyChangedEventArgs e )
    {
      var newValue = (double)e.NewValue ;

      if ( false == _textToValue ) {
        CoerceTextBoxText( HasValidValue, newValue ) ;
      }

      RaiseEvent( new ValueChangedEventArgs( ValueChangedEvent, this, (double)e.OldValue, newValue ) ) ;
    }

    private void CoerceTextBoxText( bool hasValidValue, double newValue )
    {
      if ( null == _textBox ) return ;

      _textBox.Text = ( hasValidValue ? newValue.ToString() : string.Empty ) ;
    }

    private double CoerceValueProperty( double value )
    {
      if ( _textToValue ) return value ;

      return Math.Max( MinValue, Math.Min( value, MaxValue ) ) ;
    }

    private void OnMinValueChanged( DependencyPropertyChangedEventArgs e )
    {
      RaiseEvent( new ValueChangedEventArgs( MinValueChangedEvent, this, (double)e.OldValue, (double)e.NewValue ) ) ;
    }

    private void OnMaxValueChanged( DependencyPropertyChangedEventArgs e )
    {
      RaiseEvent( new ValueChangedEventArgs( MaxValueChangedEvent, this, (double)e.OldValue, (double)e.NewValue ) ) ;
    }

    private void SelectAll()
    {
      if ( _textBox is { } textBox ) {
        this.Dispatcher.InvokeAsync( () =>
        {
          Task.Delay( 0 ) ;
          textBox.SelectAll() ;
        } ) ;
      }
    }

    private void TextBox_PreviewKeyDown( object sender, KeyEventArgs e )
    {
      if ( e.Key == Key.Down ) {
        if ( null != _down ) {
          _down.RaiseEvent( new RoutedEventArgs( ButtonBase.ClickEvent ) ) ;
          e.Handled = true ;
        }
      }
      else if ( e.Key == Key.Up ) {
        if ( null != _up ) {
          _up.RaiseEvent( new RoutedEventArgs( ButtonBase.ClickEvent ) ) ;
          e.Handled = true ;
        }
      }
    }

    private void TextBox_TextChanged( object sender, TextChangedEventArgs e )
    {
      if ( null == _textBox ) return ;

      e.Handled = true ;

      _textToValue = true ;
      try {
        HasValidValue = double.TryParse( _textBox.Text.Trim(), out var result ) || ( false == CanHaveNull ) ;
        Value = HasValidValue ? result : 0.0 ;
      }
      finally {
        _textToValue = false ;
      }
    }

    private void NumericUpDownEx_OnLostFocus( object sender, RoutedEventArgs e )
    {
      var orgValue = Value ;
      var newValue = CoerceValueProperty( orgValue ) ;
      if ( orgValue == newValue ) {
        CoerceTextBoxText( HasValidValue, newValue ) ;
      }
      else {
        Value = CoerceValueProperty( Value ) ;
      }
    }
  }
}