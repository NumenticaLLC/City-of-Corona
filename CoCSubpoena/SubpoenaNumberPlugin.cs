using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CoCSubpoena {
    public class SubpoenaNumberPlugin : IPlugin {
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
                Entity this_subpoena = (Entity)context.InputParameters["Target"];

                // Obtain the organization service reference which you will need for  
                // web service calls.  
                IOrganizationServiceFactory serviceFactory =
                    (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                try {
                    // Plug-in business logic goes here.
                    // when running for "Create"
                    if (context.MessageName == "Create") {
                        // handle case where entity passes an existing name during creation
                        if (this_subpoena.Contains("coc_name")) {
                            // 0, 4 - year
                            // 4, 1 - hyphen
                            // 5 - int -> sequence number
                            String given_name = this_subpoena["coc_name"].ToString();

                            int year, seq;
                            // String hyph = given_name.Substring(4, 1);

                            if (Int32.TryParse(given_name.Substring(0, 4), out year) && Int32.TryParse(given_name.Substring(5), out seq)) {
                                // update attributes
                                // sequence number is the autonumber
                                this_subpoena["coc_sequencenumber"] = Decimal.Parse(seq.ToString());
                                // supboena number is YYYY-N, where N is sequence number
                                this_subpoena["coc_subpoenanumber"] = year.ToString() + "-" + seq.ToString();
                                // subpoena name is "[subpoena number] [case name]"
                                if (this_subpoena.Contains("coc_casename")) {
                                    this_subpoena["coc_name"] = this_subpoena["coc_subpoenanumber"] + " " + this_subpoena["coc_casename"].ToString();
                                } else {
                                    this_subpoena["coc_name"] = this_subpoena["coc_subpoenanumber"];
                                }

                                // this_subpoena["createdon"] = new System.DateTime(year, System.DateTime.Now.Month, System.DateTime.Now.Day);
                                // this_subpoena["overridencreatedon"] = new System.DateTime(year, System.DateTime.Now.Month, System.DateTime.Now.Day);
                            } else {
                                throw new InvalidPluginExecutionException("Invalid name format.");
                            }

                        } else {
                            String gen_string = "";

                            // figure out year and append it to string
                            gen_string += System.DateTime.Now.Year.ToString();

                            gen_string += "-";

                            // figure out sequence number and append it to string
                            string query = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                                        <entity name='coc_subpoena'>
                                            <attribute name='coc_subpoenaid' />
                                            <attribute name='coc_name' />
                                            <attribute name='createdon' />
                                            <attribute name='coc_sequencenumber' />
                                            <attribute name='coc_casename' />
                                            <order attribute='coc_name' descending='false' />
                                            <filter type='and'>
                                                <condition attribute='createdon' operator='this-year' />
                                                <condition attribute='coc_sequencenumber' operator='not-null' />
                                            </filter>
                                        </entity>
                                    </fetch>";

                            EntityCollection results = service.RetrieveMultiple(new FetchExpression(query));

                            Decimal max = 0;
                            if (results.TotalRecordCount == 0) {
                                max = 1;
                            } else {
                                foreach (Entity subpoena in results.Entities) {
                                    if (max < Decimal.Parse(subpoena["coc_sequencenumber"].ToString())) {
                                        max = Decimal.Parse(subpoena["coc_sequencenumber"].ToString());
                                    }
                                }
                            }
                            gen_string += (Decimal.Truncate(max + 1)).ToString();

                            // update attributes
                            // sequence number is the autonumber
                            this_subpoena["coc_sequencenumber"] = (max + 1);
                            // supboena number is YYYY-N, where N is sequence number
                            this_subpoena["coc_subpoenanumber"] = gen_string;
                            // subpoena name is "[subpoena number] [case name]"
                            if (this_subpoena.Contains("coc_casename")) {
                                this_subpoena["coc_name"] = gen_string + " " + this_subpoena["coc_casename"].ToString();
                            } else {
                                this_subpoena["coc_name"] = gen_string;
                            }
                        }
                    }
                    // when running for "Update"
                    if (context.MessageName == "Update") {
                        // if image exists
                        if (context.PreEntityImages.Contains("subpoena_image") && context.PreEntityImages["subpoena_image"] is Entity) {
                            // retrieve pre-update image
                            Entity image_subpoena = (Entity)context.PreEntityImages["subpoena_image"];
                            // update subpoena name
                            this_subpoena["coc_name"] = image_subpoena["coc_subpoenanumber"].ToString() + " " + this_subpoena["coc_casename"].ToString();
                        }// else throw new InvalidPluginExecutionException("No Image Entity");
                    }
                } catch (FaultException<OrganizationServiceFault> ex) {
                    throw new InvalidPluginExecutionException("An error occurred in SubpoenaNumberPlugin.", ex);
                } catch (Exception ex) {
                    tracingService.Trace("SubpoenaNumberPlugin: {0}", ex.ToString());
                    throw;
                }
            }
        }
    }
}
