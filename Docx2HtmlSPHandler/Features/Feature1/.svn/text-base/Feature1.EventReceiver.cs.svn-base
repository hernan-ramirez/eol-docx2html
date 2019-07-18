using System;
using System.Runtime.InteropServices;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Administration;

namespace Docx2HtmlSPHandler.Features.Feature1
{
    /// <summary>
    /// Esta clase controla los eventos generados durante la activación, desactivación, instalación, desinstalación y actualización de características.
    /// </summary>
    /// <remarks>
    /// El GUID asociado a esta clase se puede usar durante el empaquetado y no se debe modificar.
    /// </remarks>

    [Guid("d0b98fe9-2277-4a4f-a19b-79cfa70138b7")]
    public class Feature1EventReceiver : SPFeatureReceiver
    {
        private const string WebConfigModificationOwner = "Docx2Html";

        public override void FeatureActivated(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                SPWebService service = SPWebService.ContentService;
                service.WebConfigModifications.Add(GetWebConfigModification());
                service.Update();
                service.ApplyWebConfigModifications();
            });
        }

        private static SPWebConfigModification GetWebConfigModification()
        {
            SPWebConfigModification mod = new SPWebConfigModification
            {
                Name = "add[@name=\"Errepar.Docx2Html\"][@path=\"*.docxhtml*\"][@verb=\"*\"][@type=\"Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler.Docx2Html, Docx2HtmlSPHandler, Version=1.0.0.0, Culture=neutral, PublicKeyToken=195b83f78e9dc1a4\"][@preCondition=\"integratedMode\"]",
                Path = "configuration/system.webServer/handlers",
                Type = SPWebConfigModification.SPWebConfigModificationType.EnsureChildNode,
                Value = "<add name=\"Errepar.Docx2Html\" path=\"*.docxhtml*\" verb=\"*\" type=\"Docx2HtmlSPHandler.Layouts.Docx2HtmlSPHandler.Docx2Html, Docx2HtmlSPHandler, Version=1.0.0.0, Culture=neutral, PublicKeyToken=195b83f78e9dc1a4\" preCondition=\"integratedMode\" />"
            };
            return mod;
        }


        public override void FeatureDeactivating(SPFeatureReceiverProperties properties)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate
            {
                SPSecurity.RunWithElevatedPrivileges(delegate
                {
                    SPWebService service = SPWebService.ContentService;
                    service.WebConfigModifications.Remove(GetWebConfigModification());
                    service.Update();
                    service.ApplyWebConfigModifications();
                });
            });
        }
    }
}
