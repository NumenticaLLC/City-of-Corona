using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CoCSubpoena.TORT {
    public class PreCreateClaimNumberPlugin : IPlugin {
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
                    String gen_string = "";

                    if (this_tort.Contains("coc_claimnumber")/* && this_tort.Contains("coc_datefiled")*/) {
                        // 0, 4 - month and year
                        // 4, 1 - hyphen
                        // 5 - int -> sequence number
                        String given_claim_number = this_tort["coc_claimnumber"].ToString();

                        int month_year, seq;
                        // String hyph = given_claim_number.Substring(4, 1);

                        if (Int32.TryParse(given_claim_number.Substring(0, 4), out month_year) && Int32.TryParse(given_claim_number.Substring(5), out seq)) {
                            // sequence number is the autonumber
                            this_tort["coc_sequencenumber"] = Decimal.Parse(seq.ToString());
                            gen_string += month_year.ToString() + "-" + seq.ToString();
                        }
                        else throw new InvalidPluginExecutionException("Invalid claim number format.");
                    }
                    else {
                        // figure out month/year and append it to string
                        gen_string += System.DateTime.Now.ToString("MM");
                        gen_string += System.DateTime.Now.ToString("yy");

                        gen_string += "-";

                        // figure out sequence number and append it to string
                        string query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                            <entity name='coc_tortclaims'>
                                                <attribute name='coc_tortclaimsid' />
                                                <attribute name='createdon' />
                                                <attribute name='coc_sequencenumber' />
                                                <attribute name='coc_name' />
                                                <attribute name='coc_lastname' />
                                                <attribute name='coc_firstname' />
                                                <order attribute='createdon' descending='false' />
                                                <filter type='and'>
                                                    <condition attribute='createdon' operator='this-month' />
                                                    <condition attribute='coc_sequencenumber' operator='not-null' />
                                                </filter>
                                            </entity>
                                        </fetch>";

                        EntityCollection results = service.RetrieveMultiple(new FetchExpression(query));

                        Decimal max = 0;
                        if (results.TotalRecordCount == 0) max = 1;
                        else {
                            foreach (Entity tort in results.Entities) {
                                if (max < Decimal.Parse(tort["coc_sequencenumber"].ToString())) {
                                    max = Decimal.Parse(tort["coc_sequencenumber"].ToString());
                                }
                            }
                        }
                        gen_string += (Decimal.Truncate(max + 1)).ToString();

                        // sequence number is the autonumber
                        this_tort["coc_sequencenumber"] = (max + 1);
                    }
                    // update attributes

                    // supboena number is YYYY-N, where N is sequence number
                    this_tort["coc_claimnumber"] = gen_string;

                    // tort claim name is "Claim No. [claim number] - [first name] [last name]"
                    String name = "";
                    if (this_tort.Contains("coc_firstname") || this_tort.Contains("coc_lastname")) {
                        name = String.Format(" -{0}{1}",
                            this_tort.Contains("coc_firstname") ? String.Format(" {0}", this_tort["coc_firstname"]) : "",
                            this_tort.Contains("coc_lastname") ? String.Format(" {0}", this_tort["coc_lastname"]) : "");
                    }
                    this_tort["coc_name"] = $"Claim No. {gen_string}{name}";
                }
                catch (FaultException<OrganizationServiceFault> ex) {
                    throw new InvalidPluginExecutionException("An error occurred in PreCreateTortClaimNumberPlugin.", ex);
                }
                catch (Exception ex) {
                    tracingService.Trace("PreCreateTortClaimNumberPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
