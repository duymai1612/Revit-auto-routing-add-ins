#if DUMP_LOGS
using System ;
using System.Collections.Generic ;
using System.Globalization ;
using System.Linq ;
using System.Xml ;
using Arent3d.Architecture.Routing.CollisionTree ;
using Arent3d.CollisionLib ;
using Arent3d.GeometryLib ;
using Arent3d.Routing ;
using Arent3d.Routing.Conditions ;
using Arent3d.Utility ;
using Autodesk.Revit.DB ;
using MathLib ;

namespace Arent3d.Architecture.Routing
{
  /// <summary>
  /// AutoRoutingTarget Dumper (for debug).
  /// </summary>
  internal static class LogDumper
  {
    #region Routing Targets Writer

    public static void DumpRoutingTargets( this IEnumerable<IAutoRoutingTarget> routingTargets, string file, ICollisionCheck? collision )
    {
      using var writer = XmlWriter.Create( file, new XmlWriterSettings { Indent = true, IndentChars = "  ", } ) ;

      writer.WriteStartDocument() ;
      using ( writer.WriteElement( "Routing" ) ) {
        writer.WriteTargets( "Targets", routingTargets ) ;
        if ( collision is CollisionTree.CollisionTree tree ) {
          writer.WriteCollision( "Collision", tree ) ;
        }
      }
      writer.WriteEndDocument() ;
    }

    private static void WriteTargets( this XmlWriter writer, string elmName, IEnumerable<IAutoRoutingTarget> routingTargets )
    {
      using var _ = writer.WriteElement( elmName ) ;

      foreach ( var target in routingTargets ) {
        writer.WriteTarget( "Target", target ) ;
      }
    }

    private static void WriteTarget( this XmlWriter writer, string elmName, IAutoRoutingTarget target )
    {
      using var _ = writer.WriteElement( elmName ) ;

      writer.WriteCondition( "Condition", target.Condition ) ;

      if ( target.CreateConstraints() is {} constraints ) {
        writer.WriteConstraint( "Constraints", constraints ) ;
      }
    }

    private static void WriteCondition( this XmlWriter writer, string elmName, ICommonRoutingCondition targetCondition )
    {
      using var _ = writer.WriteElement( elmName ) ;

      writer.WriteElementString( "IsRoutingOnPipeRacks", targetCondition.IsRoutingOnPipeRacks.ToString() ) ;
      writer.WriteElementString( "Type", targetCondition.Type.ToString() ) ;
      writer.WriteElementString( "Priority", targetCondition.Priority.ToString() ) ;
      writer.WriteElementString( "LoopType", targetCondition.LoopType.ToString() ) ;
      writer.WriteElementString( "AllowHorizontalBranches", targetCondition.AllowHorizontalBranches.ToString() ) ;
      if ( targetCondition.FixedBopHeight.HasValue ) {
        writer.WriteElementString( "FixedBopHeight", targetCondition.FixedBopHeight.Value.ToString( CultureInfo.InvariantCulture ) ) ;
      }
    }

    private static void WriteConstraint( this XmlWriter writer, string elmName, IAutoRoutingSpatialConstraints constraints )
    {
      using var _ = writer.WriteElement( elmName ) ;

      foreach ( var ep in constraints.Starts ) {
        writer.WriteEndPoint( "Start", ep ) ;
      }
      foreach ( var ep in constraints.Destination ) {
        writer.WriteEndPoint( "Destination", ep ) ;
      }
    }

    private static void WriteEndPoint( this XmlWriter writer, string elmName, IAutoRoutingEndPoint ep )
    {
      using var _ = writer.WriteElement( elmName ) ;

      writer.WriteElementString( "IsStart", ep.IsStart.ToString() ) ;
      writer.WriteElementString( "Depth", ep.Depth.ToString() ) ;
      writer.WriteVector( "Position", ep.Position ) ;
      writer.WriteVector( "Direction", ep.Direction ) ;
      writer.WriteElementString( "AllowHorizontalBranches", ep.AllowHorizontalBranches.ToString() ) ;
      writer.WriteElementString( "AllowThroughBatteryLimit", ep.AllowThroughBatteryLimit.ToString() ) ;
      writer.WritePipeCondition( "PipeCondition", ep.PipeCondition ) ;
    }

    private static void WritePipeCondition( this XmlWriter writer, string elmName, IRouteCondition condition )
    {
      using var _ = writer.WriteElement( elmName ) ;

     writer.WriteElementString( "Diameter", condition.Diameter.Outside.ToString( CultureInfo.InvariantCulture ) ) ;
     writer.WriteElementString( "DiameterPipeAndInsulation", condition.DiameterPipeAndInsulation.ToString( CultureInfo.InvariantCulture ) ) ;
     writer.WriteElementString( "DiameterFlangeAndInsulation", condition.DiameterFlangeAndInsulation.ToString( CultureInfo.InvariantCulture ) ) ;
     writer.WriteElementString( "ProcessConstraint", condition.ProcessConstraint.ToString() ) ;
     writer.WriteElementString( "FluidPhase", condition.FluidPhase ) ;
    }

    private static void WriteCollision( this XmlWriter writer, string elmName, CollisionTree.CollisionTree collision )
    {
      using var _ = writer.WriteElement( elmName ) ;

      foreach ( var (elmId, triangle) in collision.GetTriangles() ) {
        writer.WriteTriangleWithId( "Element", triangle, elmId ) ;
      }
    }

    private static void WriteTriangleWithId( this XmlWriter writer, string elmName, MeshTriangle triangle, ElementId id )
    {
      using var _ = writer.WriteElement( elmName ) ;
      writer.WriteAttributeString( "ElementId", id.IntegerValue.ToString() ) ;

      writer.WriteVector( "P0", triangle.get_Vertex( 0 ).To3dRaw() ) ;
      writer.WriteVector( "P1", triangle.get_Vertex( 1 ).To3dRaw() ) ;
      writer.WriteVector( "P2", triangle.get_Vertex( 2 ).To3dRaw() ) ;
    }

    private static void WriteVector( this XmlWriter writer, string elmName, Vector3d dir )
    {
      using var _ = writer.WriteElement( elmName ) ;

      writer.WriteElementString( "X", dir.x.ToString( CultureInfo.InvariantCulture ) ) ;
      writer.WriteElementString( "Y", dir.y.ToString( CultureInfo.InvariantCulture ) ) ;
      writer.WriteElementString( "Z", dir.z.ToString( CultureInfo.InvariantCulture ) ) ;
    }

    #endregion

    #region Routing Targets Reader

    public static (IReadOnlyCollection<IAutoRoutingTarget>, ICollisionCheck?) RoutingTargetsFromDump( string file )
    {
      using var reader = XmlReader.Create( file ) ;

      reader.ReadToFollowing( "Routing" ) ;
      reader.ReadStartElement( "Routing" ) ;
      var targets = reader.ReadTargets( "Targets" ) ;
      var tree = reader.ReadCollisionTree( "Collision" ) ;
      reader.ReadEndElement() ;

      return ( targets, tree ) ;
    }

    private static IReadOnlyCollection<IAutoRoutingTarget> ReadTargets( this XmlReader reader, string elmName )
    {
      var list = new List<IAutoRoutingTarget>() ;

      reader.ReadToFollowing( elmName ) ;
      reader.ReadStartElement( elmName ) ;
      while ( reader.IsStartElement() ) {
        reader.ReadStartElement( "Target" ) ;
        list.Add( ReadTarget( reader ) ) ;
        reader.ReadEndElement() ;
      }

      reader.ReadEndElement() ;

      return list ;
    }

    private class DumpedAutoRoutingTarget : IAutoRoutingTarget
    {
      public IAutoRoutingSpatialConstraints? Constraints { get ; set ; }
      public IAutoRoutingSpatialConstraints? CreateConstraints() => Constraints ;

      public ICommonRoutingCondition? Condition { get ; set ; }

      public int RouteCount => Constraints!.Starts.Count() + Constraints.Destination.Count() - 1 ;

      public Action<IEnumerable<(IAutoRoutingEndPoint, Vector3d)>> PositionInitialized => x => { } ;
    }
    private static IAutoRoutingTarget ReadTarget( XmlReader reader )
    {
      var target = new DumpedAutoRoutingTarget() ;

      reader.ReadStartElement( "Condition" ) ;
      target.Condition = ReadCondition( reader ) ;
      reader.ReadEndElement() ;

      reader.ReadStartElement( "Constraints" ) ;
      target.Constraints = ReadConstraints( reader ) ;
      reader.ReadEndElement() ;

      return target ;
    }

    private class DumpedCommonRoutingCondition : ICommonRoutingCondition
    {
      public bool IsRoutingOnPipeRacks { get ; set ; }
      public bool IsCrossingPipeRacks => false ;
      public bool IsRouteMergeEnabled => false ;
      public LineType Type { get ; set ; }
      public int Priority { get ; set ; }
      public LoopType LoopType { get ; set ; }
      public bool AllowHorizontalBranches { get ; set ; }
      public double? FixedBopHeight { get ; set ; }
    }
    private static ICommonRoutingCondition ReadCondition( XmlReader reader )
    {
      var condition = new DumpedCommonRoutingCondition() ;

      condition.IsRoutingOnPipeRacks = ReadBool( reader, "IsRoutingOnPipeRacks" ) ;
      condition.Type = ReadEnum<LineType>( reader, "Type" ) ;
      condition.Priority = ReadInt( reader, "Priority" ) ;
      condition.LoopType = ReadEnum<LoopType>( reader, "LoopType" ) ;
      condition.AllowHorizontalBranches = ReadBool( reader, "AllowHorizontalBranches" ) ;
      if ( reader.IsStartElement() ) {
        condition.FixedBopHeight = ReadDouble( reader, "FixedBopHeight" ) ;
      }

      return condition ;
    }

    private class DumpedAutoRoutingSpatialConstraints : IAutoRoutingSpatialConstraints
    {
      public IList<IAutoRoutingEndPoint> Starts { get ; } = new List<IAutoRoutingEndPoint>() ;
      public IList<IAutoRoutingEndPoint> Destination { get ; } = new List<IAutoRoutingEndPoint>() ;

      IEnumerable<IAutoRoutingEndPoint> IAutoRoutingSpatialConstraints.Starts => Starts ;
      IEnumerable<IAutoRoutingEndPoint> IAutoRoutingSpatialConstraints.Destination => Destination ;
    }
    private static IAutoRoutingSpatialConstraints ReadConstraints( XmlReader reader )
    {
      var constraints = new DumpedAutoRoutingSpatialConstraints() ;

      while ( reader.IsStartElement() ) {
        if ( "Start" == reader.Name ) {
          reader.ReadStartElement() ;
          constraints.Starts.Add( ReadEndPoint( reader ) ) ;
          reader.ReadEndElement() ;
        }
        else if ( "Destination" == reader.Name ) {
          reader.ReadStartElement() ;
          constraints.Destination.Add( ReadEndPoint( reader ) ) ;
          reader.ReadEndElement() ;
        }
      }

      return constraints ;
    }

    private class DumpedAutoRoutingEndPoint : IAutoRoutingEndPoint
    {
      public Vector3d Position { get ; set ; }
      public Vector3d Direction { get ; set ; }
      public Vector3d? PositionConstraint => null ;
      public IRouteCondition? PipeCondition { get ; set ; }
      public bool IsStart { get ; set ; }
      public int Depth { get ; set ; }
      public RoutingPointType PointType { get ; set ; }
      public ILayerStack? LinkedRack => null ;
      public bool AllowHorizontalBranches { get ; set ; }
      public bool IsDirect => false ;
      public bool AllowThroughBatteryLimit { get ; set ; }
    }
    private static IAutoRoutingEndPoint ReadEndPoint( XmlReader reader )
    {
      var ep = new DumpedAutoRoutingEndPoint() ;

      ep.IsStart = ReadBool( reader, "IsStart" ) ;
      ep.Depth = ReadInt( reader, "Depth" ) ;
      ep.Position = ReadVector( reader, "Position" ) ;
      ep.Direction = ReadVector( reader, "Direction" ) ;
      ep.PointType = ReadEnum<RoutingPointType>( reader, "PointType" ) ;
      ep.AllowHorizontalBranches = ReadBool( reader, "AllowHorizontalBranches" ) ;
      ep.AllowThroughBatteryLimit = ReadBool( reader, "AllowThroughBatteryLimit" ) ;

      reader.ReadStartElement( "PipeCondition" ) ;
      ep.PipeCondition = ReadPipeCondition( reader ) ;
      reader.ReadEndElement() ;

      return ep ;
    }

    private class DumpedRouteCondition : IRouteCondition
    {
      public IPipeDiameter? Diameter { get ; set ; }
      public double DiameterPipeAndInsulation { get ; set ; }
      public double DiameterFlangeAndInsulation { get ; set ; }
      public IPipeSpec Spec => DefaultPipeSpec.Instance ;
      public ProcessConstraint ProcessConstraint { get ; set ; }
      public string FluidPhase { get ; set ; } = string.Empty ;
    }
    private class PipeDiameter : IPipeDiameter
    {
      public PipeDiameter( double value )
      {
        Outside = value ;
        NPSmm = (int) Math.Floor( value * 1000 ) ;
      }

      public double Outside { get ; }
      public int NPSmm { get ; }
    }
    private static IRouteCondition ReadPipeCondition( XmlReader reader )
    {
      var condition = new DumpedRouteCondition() ;

      condition.Diameter = new PipeDiameter( ReadDouble( reader, "Diameter" ) ) ;
      condition.DiameterPipeAndInsulation = ReadDouble( reader, "DiameterPipeAndInsulation" ) ;
      condition.DiameterFlangeAndInsulation = ReadDouble( reader, "DiameterFlangeAndInsulation" ) ;
      condition.ProcessConstraint = ReadEnum<ProcessConstraint>( reader, "ProcessConstraint" ) ;
      condition.FluidPhase = reader.ReadElementString( "FluidPhase" ) ;

      return condition ;
    }

    private class DumpedCollisionTree : ICollisionCheck
    {
      private readonly ITree _treeBody ;

      public DumpedCollisionTree( IReadOnlyCollection<(ElementId, (Vector3d, Vector3d, Vector3d))> triangles )
      {
        var dic = new Dictionary<int, TrianglesGeometry>() ;
        foreach ( var (elementId, triangle) in triangles ) {
          if ( false == dic.TryGetValue( elementId.IntegerValue, out var geom ) ) {
            geom = new TrianglesGeometry( elementId ) ;
            dic.Add( elementId.IntegerValue, geom ) ;
          }
          geom.AddTriangle( triangle ) ;
        }

        _treeBody = CreateTreeByFactory( dic.Values.ConvertAll( ToTreeElement ) ) ;
      }

      private static TreeElement ToTreeElement( TrianglesGeometry geom )
      {
        return new TreeElement( new TriangleGeometryBody( geom ) ) ;
      }

      private static ITree CreateTreeByFactory( IReadOnlyCollection<TreeElement> treeElements )
      {
        if ( 0 == treeElements.Count ) {
          return TreeFactory.GetTreeInstanceToBuild( TreeFactory.TreeType.Dummy, null! ) ; // Dummyの場合はtreeElementsを使用しない
        }
        else {
          var tree = TreeFactory.GetTreeInstanceToBuild( TreeFactory.TreeType.Bvh, treeElements ) ;
          tree.Build() ;
          return tree ;
        }
      }

      public IEnumerable<Box3d> GetCollidedBoxes( Box3d box )
      {
        return this._treeBody.BoxIntersects( GetGeometryBodyBox( box ) ).Select( element => element.GlobalBox3d ) ;
      }

      public IEnumerable<(Box3d, IRouteCondition?, bool)> GetCollidedBoxesInDetailToRack( Box3d box )
      {
        // Aggregated Tree から呼ぶこと、これを単独で呼ばないこと
        var tuples = this._treeBody.GetIntersectsInDetailToRack( GetGeometryBodyBox( box ) ) ;
        foreach ( var tuple in tuples ) {
          yield return ( tuple.body.GetBounds(), tuple.cond, true ) ;
        }
      }

      public IEnumerable<(Box3d, IRouteCondition?, bool)> GetCollidedBoxesAndConditions( in Box3d box, CollisionCheckStructureOption option = CollisionCheckStructureOption.CheckAll )
      {
        return GetCollidedBoxesAndConditionsImpl( box, option ) ;

        IEnumerable<(Box3d, IRouteCondition?, bool)> GetCollidedBoxesAndConditionsImpl( Box3d box, CollisionCheckStructureOption option )
        {
          var tuples = this._treeBody.GetIntersectAndRoutingCondition( GetGeometryBodyBox( box ) ) ;
          foreach ( var (body, cond, isStructure) in tuples ) {
            if ( null != cond ) {
              yield return ( body.GetGlobalGeometryBox(), cond, false ) ;
              continue ;
            }

            foreach ( var geo in body.GetGlobalGeometries() ) {
              yield return ( geo.GetBounds(), null, isStructure ) ;
            }
          }
        }
      }

      private static IGeometryBody GetGeometryBodyBox( Box3d box ) => new CollisionBox( box ) ;
    }
    private static ICollisionCheck? ReadCollisionTree( this XmlReader reader, string elmName )
    {
      if ( ! reader.IsStartElement( elmName ) ) return null ;

      var list = new List<(ElementId, (Vector3d, Vector3d, Vector3d))>() ;

      reader.ReadToFollowing( elmName ) ;
      reader.ReadStartElement( elmName ) ;
      while ( reader.IsStartElement() ) {
        list.Add( ReadTriangleWithElementId( reader, "Element" ) ) ;
      }

      reader.ReadEndElement() ;

      return new DumpedCollisionTree( list ) ;
    }

    private static bool ReadBool( XmlReader reader, string tagName )
    {
      var str = reader.ReadElementString( tagName ) ;
      return bool.TryParse( str, out var val ) ? val : default ;
    }
    private static int ReadInt( XmlReader reader, string tagName )
    {
      var str = reader.ReadElementString( tagName ) ;
      return int.TryParse( str, out var val ) ? val : default ;
    }
    private static double ReadDouble( XmlReader reader, string tagName )
    {
      var str = reader.ReadElementString( tagName ) ;
      return double.TryParse( str, NumberStyles.Any, CultureInfo.InvariantCulture, out var val ) ? val : default ;
    }
    private static TEnum ReadEnum<TEnum>( XmlReader reader, string tagName ) where TEnum : struct, Enum
    {
      var str = reader.ReadElementString( tagName ) ;
      return Enum.TryParse( str, out TEnum val ) ? val : default ;
    }
    private static (ElementId, (Vector3d, Vector3d, Vector3d)) ReadTriangleWithElementId( XmlReader reader, string tagName )
    {
      reader.ReadStartElement( tagName ) ;
      var elmId = GetElementId( reader.GetAttribute( "ElementId" ) ) ;
      var p0 = ReadVector( reader, "P0" ) ;
      var p1 = ReadVector( reader, "P1" ) ;
      var p2 = ReadVector( reader, "P2" ) ;
      reader.ReadEndElement() ;

      return ( elmId, ( p0, p1, p2 ) ) ;
    }

    private static ElementId GetElementId( string? attr )
    {
      if ( null == attr || false == int.TryParse( attr, out var id ) || id < 0 ) return ElementId.InvalidElementId ;
      return new ElementId( id ) ;
    }

    private static Vector3d ReadVector( XmlReader reader, string tagName )
    {
      reader.ReadStartElement( tagName ) ;
      var x = ReadDouble( reader, "X" ) ;
      var y = ReadDouble( reader, "Y" ) ;
      var z = ReadDouble( reader, "Z" ) ;
      reader.ReadEndElement() ;

      return new Vector3d( x, y, z ) ;
    }

    private class DefaultPipeSpec : IPipeSpec
    {
      public static DefaultPipeSpec Instance { get ; } = new DefaultPipeSpec() ;
      
      public double GetLongElbowSize( IPipeDiameter diameter ) => Get90ElbowSize( diameter.Outside ) ;

      public double Get45ElbowSize( IPipeDiameter diameter ) => Get45ElbowSize( diameter.Outside ) ;

      public double GetTeeBranchLength( IPipeDiameter header, IPipeDiameter branch ) => GetTeeBranchLength( header.Outside, branch.Outside ) ;

      public double GetTeeHeaderLength( IPipeDiameter header, IPipeDiameter branch ) => GetTeeHeaderLength( header.Outside, branch.Outside ) ;

      public double GetWeldMinDistance( IPipeDiameter diameter ) => GetWeldMinDistance( diameter.Outside ) ;
      public double GetConcentricReducerCombinationLength( IPipeDiameter largeDiameter, IPipeDiameter smallDiameter ) => GetReducerLength( largeDiameter.Outside, smallDiameter.Outside ) ;

      public double GetEccentricReducerCombinationLength( IPipeDiameter largeDiameter, IPipeDiameter smallDiameter ) => GetReducerLength( largeDiameter.Outside, smallDiameter.Outside ) ;

      public string Name => "default" ;

      private DefaultPipeSpec()
      {
      }

      private double Get90ElbowSize( double diameter )
      {
        return diameter * 1.5 ;
      }

      private double Get45ElbowSize( double diameter )
      {
        return diameter * 1.5 ;
      }

      private double GetTeeHeaderLength( double headerDiameter, double branchDiameter )
      {
        if ( headerDiameter < branchDiameter ) {
          return headerDiameter * 1.0 ;
        }
        else {
          return headerDiameter * 0.5 + branchDiameter * 0.5 ;
        }
      }

      private double GetTeeBranchLength( double headerDiameter, double branchDiameter )
      {
        if ( headerDiameter < branchDiameter ) {
          return headerDiameter * 1.0 + GetReducerLength( headerDiameter, branchDiameter ) ; // provisional
        }
        else {
          return headerDiameter * 0.5 + branchDiameter * 0.5 ; // provisional
        }
      }

      private double GetReducerLength( double diameter1, double diameter2 )
      {
        return 0 ;
      }

      private double GetWeldMinDistance( double diameter )
      {
        return MEPSystemPipeSpec.MinimumShortCurveLength ;  // 1/10 inches.
      }
    }

    private class TrianglesGeometry : IGeometry, ITriangleMesh
    {
      public ElementId ElementId { get ; }
      public TrianglesGeometry( ElementId elementId )
      {
        ElementId = elementId ;
        CodSys = LocalCodSys3d.Identity ;
        _triangles = new List<Vector3d>() ;
      }

      private TrianglesGeometry( LocalCodSys3d newCodSys, TrianglesGeometry baseGeometry )
      {
        ElementId = baseGeometry.ElementId ;
        CodSys = newCodSys ;
        _triangles = baseGeometry._triangles ;
      }

      private Box3d? _bounds = null ;
      public Box3d GetBounds() => _bounds ??= CalcBounds() ;

      private Box3d CalcBounds()
      {
        double minX = double.PositiveInfinity, minY = double.PositiveInfinity, minZ = double.PositiveInfinity ;
        double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity, maxZ = double.NegativeInfinity ;

        foreach ( var (x, y, z) in Vertices ) {
          if ( x < minX ) minX = x ;
          if ( maxX < x ) maxX = x ;
          if ( y < minY ) minY = y ;
          if ( maxY < y ) maxY = y ;
          if ( z < minZ ) minZ = z ;
          if ( maxZ < z ) maxZ = z ;
        }

        if ( minX > maxX ) return Box3d.Null ;
        return new Box3d( new Vector3d( minX, minY, minZ ), new Vector3d( maxX, maxY, maxZ ) ) ;
      }

      public ITriangleMesh GetMesh( double meshingTolerance ) => this ;

      public IGeometry Localize( LocalCodSys3d codsys ) => new TrianglesGeometry( codsys.LocalizeCodSys( CodSys ), this ) ;

      public IGeometry Globalize( LocalCodSys3d codsys ) => new TrianglesGeometry( codsys.GlobalizeCodSys( CodSys ), this ) ;

      public LocalCodSys3d CodSys { get ; }

      private readonly List<Vector3d> _triangles ;
      

      public void AddTriangle( (Vector3d, Vector3d, Vector3d) triangle )
      {
        _triangles.Add( triangle.Item1 ) ;
        _triangles.Add( triangle.Item2 ) ;
        _triangles.Add( triangle.Item3 ) ;
        _asArray = null ;
      }

      private Vector3d[]? _asArray = null ;
      private Vector3d[] VertexArray => _asArray ??= Array.ConvertAll( _triangles.ToArray(), v => CodSys.LocalizePoint( v ) ) ;
      
      public ref readonly Vector3d GetVertex( int triangleIndex, int vertexIndex ) => ref VertexArray[ triangleIndex * 3 + vertexIndex ] ;

      public int Count => _triangles.Count / 3 ;
      public IEnumerable<Vector3d> Vertices => VertexArray ;

      public ref readonly Vector3d this[ int triangleIndex, int vertexIndex ] => ref VertexArray[ triangleIndex * 3 + vertexIndex ] ;
    }
    private class TriangleGeometryBody : IGeometryBody
    {
      private readonly IReadOnlyList<TrianglesGeometry> _geoms ;

      public TriangleGeometryBody( TrianglesGeometry geom )
      {
        _geoms = new[] { geom } ;
      }

      public IReadOnlyCollection<IGeometry> GetGeometries() => _geoms ;
      public IReadOnlyCollection<IGeometry> GetGlobalGeometries() => _geoms ;
      public Box3d GetGlobalGeometryBox() => _geoms[ 0 ].GetBounds() ;
    }
    
    #endregion
  }
}
#endif