using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Linq;

namespace DynamicsCrm
{
    class Program
    {
        static void Main(string[] args)
        {
            IOrganizationService oServiceProxy;
            try
            {
                //Create the Dynamics 365 Connection:
                CrmServiceClient oMSCRMConn = new CrmServiceClient("AuthType=Office365;Username=2016-00190@students.ssau.ru;" +
                    "Password=Sofia2004;URL=https://students1.crm4.dynamics.com;");

                //Create the IOrganizationService:
                oServiceProxy = (IOrganizationService)oMSCRMConn.OrganizationWebProxyClient != null ?
                    (IOrganizationService)oMSCRMConn.OrganizationWebProxyClient : (IOrganizationService)oMSCRMConn.OrganizationServiceProxy;

                if (oServiceProxy != null)
                {
                    //Get the current user ID:
                    Guid userid = ((WhoAmIResponse)oServiceProxy.Execute(new WhoAmIRequest())).UserId;

                    if (userid != Guid.Empty)
                    {
                        Console.WriteLine("Connection Successful!");
                        //Display the CRM version number and org name that you are connected to
                        Console.WriteLine("(Version: {0}; Org: {1}",
                        oMSCRMConn.ConnectedOrgVersion, oMSCRMConn.ConnectedOrgUniqueName);
                    }
                }
                else
                {
                    Console.WriteLine("Connection failed...");
                }

                var query = new QueryExpression("contact")
                {
                    TopCount = 15,
                    ColumnSet = new ColumnSet("fullname")
                };

                EntityCollection results = oServiceProxy.RetrieveMultiple(query);

                Console.WriteLine("\n LIST OF ALL CONTACTS:");
                results.Entities.ToList().ForEach(x =>
                {
                    Console.WriteLine(x.Attributes["fullname"]);
                });

                Console.WriteLine("\n LIST OF ALL ACCOUNTS:");
                var query2 = new QueryExpression("account");
                query2.ColumnSet.AllColumns = true;

                EntityCollection results2 = oServiceProxy.RetrieveMultiple(query2);

                results2.Entities.ToList().ForEach(x =>
                {
                    Console.WriteLine(x.Attributes["name"]);
                });
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error - " + ex.ToString());
            }
            Console.ReadKey();
        }

    }
}
