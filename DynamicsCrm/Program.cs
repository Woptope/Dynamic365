using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
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

                var query7 = new QueryExpression("account")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        Conditions = { new ConditionExpression("name", ConditionOperator.Equal, "TestOrganization")
                                     }
                    }

                };
                EntityCollection results7 = oServiceProxy.RetrieveMultiple(query7);


                var query0 = new QueryExpression("connection")
                {
                    ColumnSet = new ColumnSet(true),
                    Criteria = new FilterExpression
                    {
                        //FilterOperator = LogicalOperator.Or,
                        Conditions = { new ConditionExpression("record1id", ConditionOperator.Equal, new Guid("011b5966-0dc5-ea11-a812-000d3ab416a6"))
                                     }
                    }

                };

                EntityCollection results0 = oServiceProxy.RetrieveMultiple(query0);

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
                    ColumnSet = new ColumnSet("activityid"),
                    Criteria = new FilterExpression
                    {
                        FilterOperator = LogicalOperator.And,
                        Conditions = { new ConditionExpression("sender", ConditionOperator.Equal, "vladimir.spider@ssauru.onmicrosoft.com"),
                                       new ConditionExpression("torecipients", ConditionOperator.Equal, "sueburk@margiestravel.com;")
                                     }
                    }

                };

                EntityCollection results2 = oServiceProxy.RetrieveMultiple(query2);
                Console.WriteLine("\nLIST OF EMAIL'S ID:");
                results2.Entities.ToList().ForEach(x =>
                {
                    Console.WriteLine(x.Attributes["activityid"]);

                });


                //ВЫВЕСТИ ВСЕ EMAIL-Ы, В КОТОРЫХ ТЕКУЩИЙ ПОЛЬЗОВАТЕЛЬ ЛИБО ПОЛУЧАТЕЛЬ/ЛИБО ОТПРАВИТЕЛЬ ПИСЬМА
                QueryExpression query3 = new QueryExpression("email");
                query3.ColumnSet = new ColumnSet("activityid", "to", "from");

                LinkEntity EntityB = new LinkEntity("email", "activityparty", "activityid", "activityid", JoinOperator.Inner)
                {
                    Columns = new ColumnSet("activityid", "partyid", "participationtypemask"),
                    EntityAlias = "ActivityParty",
                    LinkCriteria = new FilterExpression
                    {
                        Conditions = { new ConditionExpression("participationtypemask", ConditionOperator.In, new object[] { "1", "2" }) }
                    }
                };

                LinkEntity EntityC = new LinkEntity("activityparty", "systemuser", "partyid", "systemuserid", JoinOperator.Inner)
                {
                    Columns = new ColumnSet("fullname"),
                    EntityAlias = "User"
                };
                EntityC.LinkCriteria.Conditions.Add(new ConditionExpression("fullname", ConditionOperator.Equal, "Владимир Виханов"));
                EntityB.LinkEntities.Add(EntityC);
                query3.LinkEntities.Add(EntityB);

                EntityCollection results3 = oServiceProxy.RetrieveMultiple(query3);
                Console.WriteLine("\nSECOND LIST OF EMAIL'S ID:");

                var entities = from email in results3.Entities.ToList()
                               group email by email.Attributes["activityid"] into g
                               select new { id = g.Key };

                foreach (var x in entities)
                {
                    Console.WriteLine(x.id);
                };

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
