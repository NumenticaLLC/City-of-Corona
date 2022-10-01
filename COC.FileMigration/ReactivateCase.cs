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
    public class ReactivateCase : CodeActivity
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

            int statuscode = ((OptionSetValue)eCase.Attributes["statuscode"]).Value;
            int statecode = ((OptionSetValue)eCase.Attributes["statecode"]).Value;
            SetStateRequest setStateRequest = new SetStateRequest()
            {
                EntityMoniker = new EntityReference
                {
                    Id = incident.Id,
                    LogicalName = "incident",
                },
                State = new OptionSetValue(0),
                Status = new OptionSetValue(1)
            };
            service.Execute(setStateRequest);

            Entity oCase = new Entity(eCase.LogicalName, eCase.Id);
            oCase["coc_oldstatecode"] = statecode;
            oCase["coc_oldstatuscode"] = statuscode;
            oCase["coc_wasresolved"] = true;
            service.Update(oCase);
        }
    }
}
