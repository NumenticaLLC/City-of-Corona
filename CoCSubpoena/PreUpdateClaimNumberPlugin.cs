using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CoCSubpoena.TORT {
    public class PreUpdateClaimNumberPlugin : IPlugin {
        public void Execute(IServiceProvider serviceProvider) {
            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.  
            if (context.InputParameters.Contains("Target") &&
                context.InputParameters["Target"] is Entity) {
                // Obtain the target entity from the input parameters.  
                Entity this_tort = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try {
                    // Plug-in business logic goes here.
                    // if image exists
                    if (context.PreEntityImages.Contains("tort_image") && context.PreEntityImages["tort_image"] is Entity) {
                        // retrieve pre-update image
                        Entity image_tort = (Entity)context.PreEntityImages["tort_image"];
                        // update tort claim name
                        // tort claim name is "Claim No. [claim number] - [first name] [last name]"
                        String gen_string = image_tort["coc_claimnumber"].ToString();
                        String name = "";
                        String firstname = "";
                        String lastname = "";
                        if (this_tort.Contains("coc_firstname")) {
                            if (this_tort["coc_firstname"].ToString().Length > 0) firstname = this_tort["coc_firstname"].ToString();
                        }
                        else if (image_tort.Contains("coc_firstname")) {
                            firstname = image_tort["coc_firstname"].ToString();
                        }
                        if (this_tort.Contains("coc_lastname")) {
                            if (this_tort["coc_lastname"].ToString().Length > 0) lastname = this_tort["coc_lastname"].ToString();
                        }
                        else if (image_tort.Contains("coc_lastname")) {
                            lastname = image_tort["coc_lastname"].ToString();
                        }
                        if ($"{firstname}{lastname}".Length > 0) {
                            name = String.Format(" -{0}{1}", firstname.Length > 0 ? $" {firstname}" : "", lastname.Length > 0 ? $" {lastname}" : "");
                        }
                        this_tort["coc_name"] = $"Claim No. {gen_string}{name}";
                    }// else throw new InvalidPluginExecutionException("No Image Entity");
                }
                catch (FaultException<OrganizationServiceFault> ex) {
                    throw new InvalidPluginExecutionException("An error occurred in PreUpdateTortClaimNumberPlugin.", ex);
                }
                catch (Exception ex) {
                    tracingService.Trace("PreUpdateTortClaimNumberPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
