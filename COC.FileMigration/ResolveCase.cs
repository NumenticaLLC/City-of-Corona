using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;

namespace COC.FileMigration
{
    public class ResolveCase : CodeActivity
    {
        [RequiredArgument]
        [ReferenceTarget("incident")]
        [Input("IncidentId")]
        public InArgument<EntityReference> CaseId { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            EntityReference incident = CaseId.Get<EntityReference>(executionContext);
            Entity eCase = service.Retrieve("incident", incident.Id, new ColumnSet(true));

            //int oldStateCode = eCase.GetAttributeValue<int>("coc_oldstatecode");
            int oldStatusCode = eCase.GetAttributeValue<int>("coc_oldstatuscode");


            Entity oCase = new Entity(eCase.LogicalName, eCase.Id);
            oCase["coc_oldstatecode"] = null;
            oCase["coc_oldstatuscode"] = null;
            oCase["coc_wasresolved"] = false;
            service.Update(oCase);

            Entity incidentResolution = new Entity("incidentresolution");
            incidentResolution.Attributes.Add("subject", "Problem Solved");
            incidentResolution.Attributes.Add("incidentid", new EntityReference("incident", incident.Id));

            // Close the incident with the resolution.
            var closeIncidentRequest = new CloseIncidentRequest
            {
                IncidentResolution = incidentResolution,
                Status = new OptionSetValue(oldStatusCode)
            };

            service.Execute(closeIncidentRequest);
        }
    }
}
