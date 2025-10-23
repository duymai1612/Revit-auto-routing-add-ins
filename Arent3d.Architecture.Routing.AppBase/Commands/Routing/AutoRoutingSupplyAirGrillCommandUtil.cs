using System ;
using System.Collections.Generic ;
using System.Linq ;
using Autodesk.Revit.DB ;
using Autodesk.Revit.DB.Mechanical ;
using Autodesk.Revit.UI ;
using Autodesk.Revit.UI.Selection ;

namespace Arent3d.Architecture.Routing.AppBase.Commands.Routing
{
  public static class AutoRoutingSupplyAirGrillCommandUtil
  {
    private const double Epsilon = 1e-9 ;

    public static (Connector? InConnector, List<Connector> OutConnectors) GetAirDistributionBoxConnectors(
      FamilyInstance airDistributionBox )
    {
      Connector? inConnector = null ;
      List<Connector> outConnectors = new List<Connector>() ;

      foreach ( Connector connector in airDistributionBox.MEPModel.ConnectorManager.Connectors ) {
        if ( connector == null ) continue ;
        if ( connector.Domain != Domain.DomainHvac ) continue ;

        if ( connector.Direction == FlowDirectionType.In )
          inConnector = connector ;
        else if ( connector.Direction == FlowDirectionType.Out )
          outConnectors.Add( connector ) ;
      }

      return ( inConnector, outConnectors ) ;
    }

    public static List<Connector> FindSupplyAirGrills( Document document, Connector referenceConnector )
    {
      List<Connector> grillConnectors = new List<Connector>() ;
      DuctSystemType ductSystemType = referenceConnector.DuctSystemType ;

      FilteredElementCollector collector = new FilteredElementCollector( document ).OfClass( typeof( FamilyInstance ) )
        .OfCategory( BuiltInCategory.OST_DuctTerminal ) ;

      foreach ( Element element in collector ) {
        if ( element is not FamilyInstance familyInstance ) continue ;
        if ( ! IsSupplyAirGrill( familyInstance ) ) continue ;

        foreach ( Connector connector in familyInstance.MEPModel.ConnectorManager.Connectors ) {
          if ( connector == null ) continue ;
          if ( connector.Domain != Domain.DomainHvac ) continue ;
          if ( connector.DuctSystemType == ductSystemType ) {
            grillConnectors.Add( connector ) ;
            break ;
          }
        }
      }

      return grillConnectors ;
    }

    /// <summary>
    /// Detect "Supply Air Grill" by checking DuctSystemType.
    /// Uses Revit API enum DuctSystemType.SupplyAir.
    /// </summary>
    public static bool IsSupplyAirGrill( FamilyInstance familyInstance )
    {
      foreach ( Connector connector in familyInstance.MEPModel.ConnectorManager.Connectors ) {
        if ( connector?.Domain == Domain.DomainHvac && connector.DuctSystemType == DuctSystemType.SupplyAir ) {
          return true ;
        }
      }

      return false ;
    }

    public static List<(Connector OutConnector, Connector SupplyAirGrillConnector)> CreatePairings(
      Connector inConnector, List<Connector> outConnectors, List<Connector> supplyAirGrills )
    {
      if ( inConnector == null ) throw new ArgumentNullException( nameof( inConnector ) ) ;

      Transform inCoordinateSystem = inConnector.CoordinateSystem ??
                                     throw new InvalidOperationException( "InConnector has no coordinate system." ) ;

      XYZ originIn = inCoordinateSystem.Origin ;
      XYZ basisXIn = NormalizeVector3DSafe( inCoordinateSystem.BasisX ) ;
      XYZ basisYIn = NormalizeVector3DSafe( inCoordinateSystem.BasisY ) ;
      XYZ basisZIn = NormalizeVector3DSafe( inCoordinateSystem.BasisZ ) ;

      if ( IsNearZeroVector3D( basisXIn ) || IsNearZeroVector3D( basisYIn ) || IsNearZeroVector3D( basisZIn ) )
        throw new InvalidOperationException( "InConnector CoordinateSystem is invalid." ) ;

      var outZx = PartitionByPlaneNormal( outConnectors, originIn, basisYIn ) ;
      var grillZx = PartitionByPlaneNormal( supplyAirGrills, originIn, basisYIn ) ;

      var outZy = PartitionByPlaneNormal( outConnectors, originIn, basisXIn ) ;
      var grillZy = PartitionByPlaneNormal( supplyAirGrills, originIn, basisXIn ) ;

      var scoreZx = Math.Min( outZx.Upper.Count, grillZx.Upper.Count ) +
                    Math.Min( outZx.Lower.Count, grillZx.Lower.Count ) ;
      var scoreZy = Math.Min( outZy.Upper.Count, grillZy.Upper.Count ) +
                    Math.Min( outZy.Lower.Count, grillZy.Lower.Count ) ;

      var useZx = scoreZx >= scoreZy ;
      var chosenOut = useZx ? outZx : outZy ;
      var chosenGrill = useZx ? grillZx : grillZy ;

      var upperOutSorted = SortByAngleWithZ( chosenOut.Upper, originIn, basisZIn ) ;
      var lowerOutSorted = SortByAngleWithZ( chosenOut.Lower, originIn, basisZIn ) ;
      var upperGrillSorted = SortByAngleWithZ( chosenGrill.Upper, originIn, basisZIn ) ;
      var lowerGrillSorted = SortByAngleWithZ( chosenGrill.Lower, originIn, basisZIn ) ;

      var pairings = new List<(Connector, Connector)>() ;
      int upperPair = Math.Min( upperOutSorted.Count, upperGrillSorted.Count ) ;
      for ( int i = 0 ; i < upperPair ; ++i ) pairings.Add( ( upperOutSorted[ i ], upperGrillSorted[ i ] ) ) ;

      int lowerPair = Math.Min( lowerOutSorted.Count, lowerGrillSorted.Count ) ;
      for ( int i = 0 ; i < lowerPair ; ++i ) pairings.Add( ( lowerOutSorted[ i ], lowerGrillSorted[ i ] ) ) ;

      var remainingOut = new List<Connector>() ;
      remainingOut.AddRange( upperOutSorted.Skip( upperPair ) ) ;
      remainingOut.AddRange( lowerOutSorted.Skip( lowerPair ) ) ;

      var remainingGrill = new List<Connector>() ;
      remainingGrill.AddRange( upperGrillSorted.Skip( upperPair ) ) ;
      remainingGrill.AddRange( lowerGrillSorted.Skip( lowerPair ) ) ;

      int rem = Math.Min( remainingOut.Count, remainingGrill.Count ) ;
      for ( int i = 0 ; i < rem ; ++i ) pairings.Add( ( remainingOut[ i ], remainingGrill[ i ] ) ) ;

      return pairings ;
    }

    private static (List<Connector> Upper, List<Connector> Lower) PartitionByPlaneNormal( List<Connector> connectors,
      XYZ originIn, XYZ planeNormal )
    {
      var upper = new List<Connector>() ;
      var lower = new List<Connector>() ;

      foreach ( var connector in connectors ) {
        XYZ v = connector.Origin - originIn ;
        double s = Dot( v, planeNormal ) ;
        if ( s >= -Epsilon ) upper.Add( connector ) ;
        else lower.Add( connector ) ;
      }

      return ( upper, lower ) ;
    }

    private static List<Connector> SortByAngleWithZ( List<Connector> connectors, XYZ originIn, XYZ basisZIn )
    {
      var items = new List<(Connector C, double AngleDeg, double Dist)>() ;

      foreach ( var c in connectors ) {
        XYZ v = c.Origin - originIn ;
        double d = Length( v ) ;
        XYZ vn = d < Epsilon ? XYZ.Zero : new XYZ( v.X / d, v.Y / d, v.Z / d ) ;

        double dot = Dot( vn, basisZIn ) ;
        if ( dot < -1.0 ) dot = -1.0 ;
        else if ( dot > 1.0 ) dot = 1.0 ;

        double ang = Math.Acos( dot ) * 180.0 / Math.PI ;
        items.Add( ( c, ang, d ) ) ;
      }

      return items.OrderBy( t => t.AngleDeg ).ThenBy( t => t.Dist ).Select( t => t.C ).ToList() ;
    }

    private static double Dot( XYZ a, XYZ b ) => a.X * b.X + a.Y * b.Y + a.Z * b.Z ;
    private static double Length( XYZ v ) => Math.Sqrt( v.X * v.X + v.Y * v.Y + v.Z * v.Z ) ;
    private static bool IsNearZeroVector3D( XYZ v, double eps = Epsilon ) => Length( v ) < eps ;

    private static XYZ NormalizeVector3DSafe( XYZ v )
    {
      double n = Length( v ) ;
      return n < Epsilon ? XYZ.Zero : new XYZ( v.X / n, v.Y / n, v.Z / n ) ;
    }

    public class AirDistributionBoxSelectionFilter : ISelectionFilter
    {
      public bool AllowElement( Element element )
      {
        try {
          if ( element is not FamilyInstance familyInstance ) return false ;
          Category category = element.Category ;
          if ( category == null ) return false ;

          int categoryId = category.Id.IntegerValue ;
          bool isValidCategory = categoryId == (int)BuiltInCategory.OST_DuctAccessory ||
                                 categoryId == (int)BuiltInCategory.OST_MechanicalEquipment ;
          if ( ! isValidCategory ) return false ;

          MEPModel mepModel = familyInstance.MEPModel ;
          if ( mepModel == null ) return false ;

          ConnectorManager connectorManager = mepModel.ConnectorManager ;
          if ( connectorManager == null ) return false ;

          bool hasInConnector = false ;
          bool hasOutConnector = false ;

          foreach ( Connector connector in connectorManager.Connectors ) {
            if ( connector == null ) continue ;
            if ( connector.Domain != Domain.DomainHvac ) continue ;

            if ( connector.Direction == FlowDirectionType.In ) hasInConnector = true ;
            else if ( connector.Direction == FlowDirectionType.Out ) hasOutConnector = true ;
            if ( hasInConnector && hasOutConnector ) return true ;
          }

          return false ;
        }
        catch {
          return false ;
        }
      }

      public bool AllowReference( Reference reference, XYZ position ) => false ;
    }
  }
}
