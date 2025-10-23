using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExternalService;
using Autodesk.Revit.UI;

namespace Arent3d.Architecture.Routing.AppBase.UI.ExternalGraphics
{
    public class DrawingServerHost
    {
        private readonly HashSet<Document> documentList;
        private readonly List<DrawingServer> serverList;

        public DrawingServerHost()
        {
            this.documentList = new HashSet<Document>();
            this.serverList = new List<DrawingServer>();
        }

        public void RegisterServer(DrawingServer drawingServer)
        {
            if (this.IsRegisterServer(drawingServer.Document))
            {
                return;
            }

            var directContext3DService =
                ExternalServiceRegistry.GetService(ExternalServices.BuiltInExternalServices.DirectContext3DService);
            if (directContext3DService is MultiServerService msDirectContext3DService)
            {
                var serverIds = msDirectContext3DService.GetActiveServerIds();
                directContext3DService.AddServer(drawingServer);
                this.serverList.Add(drawingServer);

                serverIds.Add(drawingServer.GetServerId());

                msDirectContext3DService.SetActiveServers(serverIds);
            }

            this.documentList.Add(drawingServer.Document);
        }

        public void UnRegisterServer(Document document)
        {
            if (!this.IsRegisterServer(document))
            {
                return;
            }

            var externalDrawerServiceId = ExternalServices.BuiltInExternalServices.DirectContext3DService;
            if (!(ExternalServiceRegistry.GetService(externalDrawerServiceId) is MultiServerService
                externalDrawerService))
            {
                return;
            }

            foreach (var registeredServerId in externalDrawerService.GetRegisteredServerIds())
            {
                if (!(externalDrawerService.GetServer(registeredServerId) is DrawingServer externalDrawServer))
                {
                    continue;
                }

                if (document != null && !document.Equals(externalDrawServer.Document))
                {
                    continue;
                }

                externalDrawerService.RemoveServer(registeredServerId);
                var uidoc = new UIDocument(document);
                uidoc.UpdateAllOpenViews();
                break;
            }
        }

        public bool IsRegisterServer(Document document)
        {
            try
            {
                var externalDrawerServiceId = ExternalServices.BuiltInExternalServices.DirectContext3DService;
                if (!(ExternalServiceRegistry.GetService(externalDrawerServiceId) is MultiServerService
                    externalDrawerService))
                {
                    return false;
                }

                foreach (var registeredServerId in externalDrawerService.GetRegisteredServerIds())
                {
                    if (!(externalDrawerService.GetServer(registeredServerId) is DrawingServer
                        externalDrawServer))
                    {
                        continue;
                    }

                    if (document != null && !document.Equals(externalDrawServer.Document))
                    {
                        continue;
                    }
                }
            }
            catch (Exception)
            {
                //
            }

            return false;
        }

    }
}