using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.DirectContext3D;
using Autodesk.Revit.DB.ExternalService;

namespace Arent3d.Architecture.Routing.AppBase.UI.ExternalGraphics
{
#pragma warning disable CS8618 // Non-nullable field must contain a non-null value when exiting constructor. Consider declaring as nullable.

    public class DrawingServer : IDirectContext3DServer
    {
        private readonly Guid _guid;

        private RenderingPassBufferStorage _edgeBufferStorage;

        public DrawingServer(Document doc)
        {
            this._guid = Guid.NewGuid();
            this.Document = doc;
        }

        public Document Document { get; set; }
        public Guid GetServerId()
        {
            return this._guid;
        }

        public string GetVendorId()
        {
            return "STSO";
        }

        public ExternalServiceId GetServiceId()
        {
            return ExternalServices.BuiltInExternalServices.DirectContext3DService;
        }

        public virtual string GetName()
        {
            return "";
        }

        public virtual string GetDescription()
        {
            return "";
        }

        public string GetApplicationId()
        {
            return "";
        }

        public string GetSourceId()
        {
            return "";
        }

        public bool UsesHandles()
        {
            return false;
        }

        public virtual bool CanExecute(View view)
        {
            var doc = view.Document;
            return doc.Equals(this.Document);
        }

        public virtual Outline? GetBoundingBox(View view)
        {
            return null;
        }

        public bool UseInTransparentPass(View view)
        {
            return true;
        }

        // Submits the geometry for rendering.
        public virtual void RenderScene(View view, DisplayStyle displayStyle)
        {
            try
            {
                // Populate geometry buffers if they are not initialized or need updating.
                this.CreateBufferStorageForElement(displayStyle, view);

                // Conditionally submit line segment primitives.
                if (displayStyle != DisplayStyle.Shading && this._edgeBufferStorage.PrimitiveCount > 0)
                {
                    DrawContext.FlushBuffer(this._edgeBufferStorage.VertexBuffer,
                        this._edgeBufferStorage.VertexBufferCount, this._edgeBufferStorage.IndexBuffer,
                        this._edgeBufferStorage.IndexBufferCount, this._edgeBufferStorage.VertexFormat,
                        this._edgeBufferStorage.EffectInstance, PrimitiveType.LineList, 0,
                        this._edgeBufferStorage.PrimitiveCount);
                }
            }
            catch (Exception)
            {
                //
            }
        }

        public virtual List<Line> PrepareProfile()
        {
            return new List<Line>();
        }

        public virtual List<Arc> ArcPrepareProfile()
        {
            return new List<Arc>() ;
        }

        public virtual void CreateBufferStorageForElement(DisplayStyle displayStyle, View view)
        {
            this._edgeBufferStorage = new RenderingPassBufferStorage(displayStyle);

            var lines = this.PrepareProfile();
            var circles = this.ArcPrepareProfile() ;
            if ( circles.Count > 0 ) {
                foreach (var edge in circles)
                {
                    var xyzs = edge.Tessellate();
                    this._edgeBufferStorage.VertexBufferCount += xyzs.Count;
                    this._edgeBufferStorage.PrimitiveCount += xyzs.Count - 1;
                    this._edgeBufferStorage.EdgeXYZs.Add(xyzs);
                }
            }
            else {
                foreach (var edge in lines)
                {
                    var xyzs = edge.Tessellate();
                    this._edgeBufferStorage.VertexBufferCount += xyzs.Count;
                    this._edgeBufferStorage.PrimitiveCount += xyzs.Count - 1;
                    this._edgeBufferStorage.EdgeXYZs.Add(xyzs);
                }
            }

            this.ProcessEdges(this._edgeBufferStorage);
        }

        private void ProcessEdges(RenderingPassBufferStorage bufferStorage)
        {
            var edges = bufferStorage.EdgeXYZs;
            if (edges.Count == 0)
            {
                return;
            }

            // Edges are encoded as line segment primitives whose vertices contain only position information.
            bufferStorage.FormatBits = VertexFormatBits.Position;

            var edgeVertexBufferSizeInFloats = VertexPosition.GetSizeInFloats() * bufferStorage.VertexBufferCount;
            var numVerticesInEdgesBefore = new List<int>();
            numVerticesInEdgesBefore.Add(0);

            bufferStorage.VertexBuffer = new VertexBuffer(edgeVertexBufferSizeInFloats);
            bufferStorage.VertexBuffer.Map(edgeVertexBufferSizeInFloats);
            {
                var vertexStream = bufferStorage.VertexBuffer.GetVertexStreamPosition();
                foreach (var xyzs in edges)
                {
                    foreach (var vertex in xyzs)
                    {
                        vertexStream.AddVertex(new VertexPosition(vertex));
                    }

                    numVerticesInEdgesBefore.Add(numVerticesInEdgesBefore.Last() + xyzs.Count);
                }
            }
            bufferStorage.VertexBuffer.Unmap();

            var edgeNumber = 0;
            bufferStorage.IndexBufferCount = bufferStorage.PrimitiveCount * IndexLine.GetSizeInShortInts();
            var indexBufferSizeInShortInts = 1 * bufferStorage.IndexBufferCount;
            bufferStorage.IndexBuffer = new IndexBuffer(indexBufferSizeInShortInts);
            bufferStorage.IndexBuffer.Map(indexBufferSizeInShortInts);
            {
                var indexStream = bufferStorage.IndexBuffer.GetIndexStreamLine();
                foreach (var xyzs in edges)
                {
                    var startIndex = numVerticesInEdgesBefore[edgeNumber];
                    for (var i = 1; i < xyzs.Count; i++)
                    {
                        // Add two indices that define a line segment.
                        indexStream.AddLine(new IndexLine(startIndex + i - 1,
                            startIndex + i));
                    }

                    edgeNumber++;
                }
            }
            bufferStorage.IndexBuffer.Unmap();

            bufferStorage.VertexFormat = new VertexFormat(bufferStorage.FormatBits);
            bufferStorage.EffectInstance = new EffectInstance(bufferStorage.FormatBits);
        }
    }
}
