using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoCSubpoena
{
    public class PreCreate_DocumentLocation_Plugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            
            #region Organization Service Consumption

            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            #endregion

            #region Entity & Attribute to Update after retrieving Sharepoint URL
            string lookForEntity = "coc_subpoena";
            string attributeToUpdate = "coc_documentlocation";
            #endregion

            if (context.InputParameters.Contains("Target") || context.InputParameters["Target"] is Entity)
            {
                Entity spdl = (Entity)context.InputParameters["Target"];
                tracingService.Trace($"Entity reference initiated with id {spdl.Id}");
                EntityReference regardingObjectId = (EntityReference)spdl.Attributes["regardingobjectid"];
                tracingService.Trace($"Regarding ObjetId is {regardingObjectId.Id} and logical name {regardingObjectId.LogicalName}");
                if (regardingObjectId.LogicalName.Equals(lookForEntity))
                {
                    tracingService.Trace($"Regarding object id is {lookForEntity} with Id {regardingObjectId.Id}");
                    string relativeUrl = spdl.Attributes.ContainsKey("relativeurl") ? spdl.GetAttributeValue<string>("relativeurl") : string.Empty;
                    tracingService.Trace($"Relative Url Found : {relativeUrl}");
                    string sharepointUrl = this.GetSharepointURL(service);
                    tracingService.Trace($"Sharepoint URL found : {sharepointUrl}");

                    if(relativeUrl != string.Empty && sharepointUrl != string.Empty)
                    {
                        tracingService.Trace("both sharepoint url and relative url are not empty.");
                        string sharepointFolderpath = string.Concat(sharepointUrl.Replace("/sites", "/:f:/r/sites"), "/", regardingObjectId.LogicalName, "/",relativeUrl);
                        tracingService.Trace($"Sharepoint Path is {sharepointFolderpath}");

                        Entity cocEntity = new Entity(lookForEntity, regardingObjectId.Id);
                        cocEntity[attributeToUpdate] = sharepointFolderpath;
                        service.Update(cocEntity);
                        tracingService.Trace($"{lookForEntity} is updated");
                    }
                }
            }
        }

        /// <summary>
        /// Returns Sharepoint URL from Sharepoint Site Entity.
        /// </summary>
        /// <param name="service">IOrganizationService object</param>
        /// <returns>String Containing Sharepoint URL</returns>
        private string GetSharepointURL(IOrganizationService service)
        {
            string query = @"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"">
                              <entity name=""sharepointsite"">
                                <attribute name=""name"" />
                                <attribute name=""parentsite"" />
                                <attribute name=""relativeurl"" />
                                <attribute name=""absoluteurl"" />
                                <attribute name=""validationstatus"" />
                                <attribute name=""isdefault"" />
                                <order attribute=""name"" descending=""false"" />
                                <filter type=""and"">
                                  <condition attribute=""servicetype"" operator=""eq"" value=""0"" />
                                  <condition attribute=""isdefault"" operator=""eq"" value=""1"" />
                                  <condition attribute=""validationstatus"" operator=""eq"" value=""4"" />
                                </filter>
                              </entity>
                            </fetch>";
            EntityCollection ecoll = service.RetrieveMultiple(new FetchExpression(query));
            return ecoll.Entities
                .Where(row => ecoll.Entities.Count > 0)
                .Select(row => row.Attributes.ContainsKey("absoluteurl") ? 
                           row.GetAttributeValue<string>("absoluteurl"): string.Empty).FirstOrDefault();
        }
    }
}
