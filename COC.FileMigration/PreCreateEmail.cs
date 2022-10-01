using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Xrm.Sdk.Query;
namespace COC.FileMigration
{
    public class PreCreateEmail : CodeActivity
    {
        [RequiredArgument]
        [ReferenceTarget("incident")]
        [Input("IncidentId")]
        public InArgument<EntityReference> CaseId { get; set; }

       
        [Output("coc_foldername")]
        public OutArgument<string> OutFolderName { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            EntityReference incident = CaseId.Get<EntityReference>(executionContext);
            Entity eCase = service.Retrieve("incident", incident.Id, new ColumnSet(true));
            string documentLocation = eCase.GetAttributeValue<string>("coc_documentlocation");
            string foldername = null;
            if (!string.IsNullOrEmpty(documentLocation))
            {
                string[] segments = documentLocation.Split('/');
                if(segments.Length == 9)
                {
                   foldername  = "/" + segments[segments.Length - 1];
                }
            }
            OutFolderName.Set(executionContext, foldername);
        }
    }
}
