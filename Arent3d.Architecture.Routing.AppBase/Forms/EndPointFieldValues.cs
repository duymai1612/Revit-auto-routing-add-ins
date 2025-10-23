using Arent3d.Architecture.Routing.EndPoints ;

namespace Arent3d.Architecture.Routing.AppBase.Forms
{
  public static class EndPointFieldValues
  {
    public static IEndPointVisitor<string> IdGetter { get ; } = new EndPointIdGetter() ;

    public static IEndPointVisitor<string> SubIdGetter { get ; } = new EndPointSubIdGetter() ;

    private class EndPointIdGetter : IEndPointVisitor<string>
    {
      public string Visit( ConnectorEndPoint endPoint ) => endPoint.EquipmentId.IntegerValue.ToString() ;
      public string Visit( RouteEndPoint endPoint ) => endPoint.RouteName ;
      public string Visit( PassPointEndPoint endPoint ) => endPoint.PassPointId.IntegerValue.ToString() ;
      public string Visit( PassPointBranchEndPoint endPoint ) => endPoint.PassPointId.IntegerValue.ToString() ;
      public string Visit( TerminatePointEndPoint endPoint ) => endPoint.TerminatePointId.IntegerValue.ToString() ;
    }

    private class EndPointSubIdGetter : IEndPointVisitor<string>
    {
      public string Visit( ConnectorEndPoint endPoint )=> endPoint.ConnectorIndex.ToString() ;
      public string Visit( RouteEndPoint endPoint ) => string.Empty ;
      public string Visit( PassPointEndPoint endPoint ) => string.Empty ;
      public string Visit( PassPointBranchEndPoint endPoint ) => string.Empty ;
      public string Visit( TerminatePointEndPoint endPoint ) => endPoint.LinkedInstanceId.IntegerValue.ToString() ;
    }
  }
}