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

                //ИЗВЛЕЧЬ ВСЕ КОНТАКТЫ, ЧЬЯ КОМПАНИЯ НАХОДИТСЯ В САМАРЕ
                var query = new QueryExpression("contact")
                {
                    ColumnSet = new ColumnSet("fullname"),
                    LinkEntities =
                         {
                             new LinkEntity
                             {
                                 JoinOperator = JoinOperator.Inner,
                                 LinkFromAttributeName = "accountid",
                                 LinkFromEntityName = "contact",
                                 LinkToAttributeName = "accountid",
                                 LinkToEntityName = "account",
                                 LinkCriteria =
                                     {
                                         Conditions =
                                         {
                                             new ConditionExpression("address1_city", ConditionOperator.Equal, "Samara")
                                         }
                                     }
                             }
                     }
                };

                EntityCollection results = oServiceProxy.RetrieveMultiple(query);

                Console.WriteLine("\nLIST OF CONTACTS, WHICH COMPANY'S CITY IS SAMARA:");
                results.Entities.ToList().ForEach(x =>
                {
                    Console.WriteLine(x.Attributes["fullname"]);
                });



                //ИЗВЛЕЧЬ ВСЕ EMAIL ЗАДАННОГО ПОЛЬЗОВАТЕЛЯ, НАПИСАННЫЕ В ОПРЕДЕЛЕННУЮ КОМПАНИЮ
                var query2 = new QueryExpression("email")
                {
                    ColumnSet = new ColumnSet("messageid"),
                    Criteria = new FilterExpression
                    {
                        FilterOperator = LogicalOperator.And,
                        Conditions = { new ConditionExpression("sender", ConditionOperator.Equal, "vladimir.spider@ssauru.onmicrosoft.com"),
                                       new ConditionExpression("torecipients", ConditionOperator.Equal, "sueburk@margiestravel.com;")
                                     }
                    }

                };

                EntityCollection results2 = oServiceProxy.RetrieveMultiple(query2);
                Console.WriteLine("\nLIST OF MAIL'S ID:");
                results2.Entities.ToList().ForEach(x =>
                {
                    Console.WriteLine(x.Attributes["messageid"]);

                });


                Console.WriteLine("ОК");
            }

            catch (Exception ex)
            {
                Console.WriteLine("Error - " + ex.ToString());
            }
            Console.ReadKey();
        }

    }
}
