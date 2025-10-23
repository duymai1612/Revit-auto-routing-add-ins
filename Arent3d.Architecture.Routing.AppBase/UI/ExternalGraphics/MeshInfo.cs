using Autodesk.Revit.DB;

namespace Arent3d.Architecture.Routing.AppBase.UI.ExternalGraphics
{
    public class MeshInfo
    {
        public readonly ColorWithTransparency ColorWithTransparency;

        public readonly Mesh Mesh;

        public readonly XYZ Normal;

        public MeshInfo(Mesh mesh, XYZ normal, ColorWithTransparency color)
        {
            this.Mesh = mesh;
            this.Normal = normal;
            this.ColorWithTransparency = color;
        }
    }
}