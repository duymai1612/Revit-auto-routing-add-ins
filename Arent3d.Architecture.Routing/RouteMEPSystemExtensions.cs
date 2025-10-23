using System ;
using System.Collections.Generic ;
using System.Linq ;
using Autodesk.Revit.DB ;
using Arent3d.Revit ;
using Arent3d.Utility ;
using Autodesk.Revit.DB.Electrical ;
using Autodesk.Revit.DB.Mechanical ;
using Autodesk.Revit.DB.Plumbing ;


namespace Arent3d.Architecture.Routing
{
  public static class RouteMEPSystemExtensions
  {
    /// <summary>
    /// Get NominalDiameterList
    /// </summary>
    /// <param name="type"></param>
    /// <param name="diameterTolerance"></param>
    /// <returns></returns>
    public static List<double> GetNominalDiameters( this MEPCurveType type, double diameterTolerance )
    {
      var resultList = new List<double>() ;
      Segment? segment = null ;
      if ( type is not ConduitType ) {
        segment = type.GetTargetSegment() ;
      }

      var diameterList = type switch
      {
        DuctType => DuctSizeSettings.GetDuctSizeSettings( type.Document )[ DuctShape.Round ].Where( s => s.UsedInSizeLists && s.UsedInSizing ).Select( s => s.NominalDiameter ).ToList(),
        PipeType => segment?.GetSizes().Where( s => s.UsedInSizeLists && s.UsedInSizing && type.HasAnyNominalDiameter( s.NominalDiameter, diameterTolerance ) ).Select( s => s.NominalDiameter ).ToList(),
        ConduitType => ConduitSizeSettings.GetConduitSizeSettings( type.Document ).Where( c => c.Key == type.get_Parameter( BuiltInParameter.CONDUIT_STANDARD_TYPE_PARAM ).AsValueString() ).Select( c => c.Value ).FirstOrDefault().Select( c => c.NominalDiameter ).ToList(),
        _ => throw new ArgumentOutOfRangeException( nameof( type ) ),
      } ;

      if ( diameterList is { } dList ) {
        resultList = dList ;
      }

      resultList.Sort() ;

      return resultList ;
    }

    public static bool HasAnyNominalDiameter( this MEPCurveType type, double nominalDiameter, double diameterTolerance )
    {
      var document = type.Document ;
      return type.RoutingPreferenceManager.GetRules( RoutingPreferenceRuleGroupType.Segments ).All( rule => HasAnyNominalDiameter( document, rule, nominalDiameter, diameterTolerance ) ) ;
    }

    private static bool HasAnyNominalDiameter( Document document, RoutingPreferenceRule rule, double nominalDiameter, double diameterTolerance )
    {
      if ( false == rule.GetCriteria().OfType<PrimarySizeCriterion>().All( criterion => criterion.IsMatchRange( nominalDiameter ) ) ) return false ;

      var segment = document.GetElementById<Segment>( rule.MEPPartId ) ;
      return ( null != segment ) && segment.HasAnyNominalDiameter( nominalDiameter, diameterTolerance ) ;
    }

    private static bool HasAnyNominalDiameter( this Segment segment, double nominalDiameter, double diameterTolerance )
    {
      return segment.GetSizes().Any( size => Math.Abs( size.NominalDiameter - nominalDiameter ) < diameterTolerance ) ;
    }

    public static IEnumerable<RoutingPreferenceRule> GetRules( this RoutingPreferenceManager rpm, RoutingPreferenceRuleGroupType groupType )
    {
      var count = rpm.GetNumberOfRules( groupType ) ;
      for ( var i = 0 ; i < count ; ++i ) {
        yield return rpm.GetRule( groupType, i ) ;
      }
    }

    public static IEnumerable<RoutingCriterionBase> GetCriteria( this RoutingPreferenceRule rule )
    {
      var count = rule.NumberOfCriteria ;
      for ( var i = 0 ; i < count ; ++i ) {
        yield return rule.GetCriterion( i ) ;
      }
    }

    public static bool IsCompatibleCurveType( this MEPCurveType curveType, Type targetCurveType )
    {
      return ( curveType.GetType() == targetCurveType ) ;
    }

    public static bool IsMatchRange( this PrimarySizeCriterion criterion, double nominalDiameter )
    {
      return criterion.MinimumSize <= nominalDiameter && nominalDiameter <= criterion.MaximumSize ;
    }

    /// <summary>
    /// Get Target SystemTypeList
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="systemClassificationInfo"></param>
    /// <returns></returns>
    public static IEnumerable<MEPSystemType> GetSystemTypes( this Document doc, MEPSystemClassificationInfo systemClassificationInfo )
    {
      return doc.GetAllElements<MEPSystemType>().Where( systemClassificationInfo.IsCompatibleTo ) ;
    }

    /// <summary>
    /// Get Conduit StandardTypeList in project
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static IEnumerable<string> GetStandardTypes( this Document doc )
    {
      return ConduitSizeSettings.GetConduitSizeSettings( doc ).Select( c => c.Key ) ;
    }


    /// <summary>
    /// Get compatible curve types.
    /// </summary>
    /// <param name="doc"></param>
    /// <param name="mepCurveTypeClass"></param>
    /// <returns></returns>
    public static IEnumerable<MEPCurveType> GetCurveTypes( this Document doc, Type? mepCurveTypeClass )
    {
      if ( null == mepCurveTypeClass ) return Enumerable.Empty<MEPCurveType>() ;

      return doc.GetAllElements<MEPCurveType>().Where( s => s.IsCompatibleCurveType( mepCurveTypeClass ) ).Select( s => s ) ;
    }

    private static Segment? GetTargetSegment( this MEPCurveType type )
    {
      if ( type.RoutingPreferenceManager is not { } manager ) return null ;

      var rules = manager.GetRules( RoutingPreferenceRuleGroupType.Segments ) ;
      var document = type.Document ;
      return rules.Select( rule => document.GetElementById<Segment>( rule.MEPPartId ) ).FirstOrDefault( segment => null != segment ) ;
    }


    private static readonly IReadOnlyDictionary<MEPSystemClassification, Domain> _classificationToDomain = GetClassificationToDomain() ;

    private static IReadOnlyDictionary<MEPSystemClassification, Domain> GetClassificationToDomain()
    {
      var dir = new Dictionary<MEPSystemClassification, Domain>() ;

      AddAllSystemClassifications<PipeSystemType>( dir, Domain.DomainPiping ) ;
      AddAllSystemClassifications<DuctSystemType>( dir, Domain.DomainHvac ) ;
      AddAllSystemClassifications<ElectricalSystemType>( dir, Domain.DomainCableTrayConduit ) ;

      return dir ;
    }

    private static void AddAllSystemClassifications<T>( Dictionary<MEPSystemClassification, Domain> dir, Domain domain ) where T : Enum
    {
      foreach ( var classification in GetAllSystemClassifications<T>() ) {
        if ( dir.ContainsKey( classification ) ) {
          dir[ classification ] = Domain.DomainUndefined ;
        }
        else {
          dir.Add( classification, domain ) ;
        }
      }
    }

    private static IEnumerable<MEPSystemClassification> GetAllSystemClassifications<T>() where T : Enum
    {
      foreach ( var name in Enum.GetNames( typeof( T ) ) ) {
        if ( false == Enum.TryParse( name, out MEPSystemClassification result ) ) continue ;

        yield return result ;
      }
    }

    public static Domain GetDomain( this MEPSystemType mepSystemType )
    {
      if ( _classificationToDomain.TryGetValue( mepSystemType.SystemClassification, out var domain ) ) return domain ;

      return Domain.DomainUndefined ;
    }
  }
}