using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;

namespace UserRoleAudit
{
    class Program
    {
        static void Main(string[] args)
        {
            CrmServiceClient service = Connect("Connect");
            if (service.IsReady)
            {
                Console.WriteLine("SuccessfullyConnected");
                EntityCollection EnabledUsers = GetUsers(service);

                //List<AuditRole> auditRoleList = new List<AuditRole>();
                List<SystemUser> userList = new List<SystemUser>();

                foreach (var user in EnabledUsers.Entities)
                {
                    SystemUser sysuser = new SystemUser();
                    sysuser.FullName= GetFullNameFromId(service, user.Id);
                    var changeRequest = new RetrieveRecordChangeHistoryRequest();
                    changeRequest.Target = new EntityReference(user.LogicalName, user.Id);

                    var changeResponse =
                        (RetrieveRecordChangeHistoryResponse)service.Execute(changeRequest);

                    AuditDetailCollection details = changeResponse.AuditDetailCollection;
                    var test = details.AuditDetails.Where(a => a.GetType().GetProperty("AccessTime") != null).OrderByDescending(a => (DateTime)a.GetType().GetProperty("AccessTime").GetValue(a)).FirstOrDefault();
                    sysuser.LastLogin = test != null ? ((DateTime)test.GetType().GetProperty("AccessTime").GetValue(test)).ToString("MM/dd/yyyy") : "";
                    userList.Add(sysuser);
                   // details.Where(a => a.<Microsoft.Crm.Sdk.Messages.UserAccessAuditDetail>());


                    //foreach (AuditDetail detail in details.AuditDetails)
                    //{

                    //    var RelationshipName = detail.GetType().GetProperty("RelationshipName") != null ? (string)detail.GetType().GetProperty("RelationshipName").GetValue(detail) : null;
                    //    if (RelationshipName != null && RelationshipName == "systemuserroles_association")
                    //    {
                    //        //if(detail.AuditRecord.GetAttributeValue<OptionSetValue>("action") == new OptionSetValue(33))
                    //        // Display some of the detail information in each audit record. 
                    //        AuditRole auditRole = DisplayAuditDetails(service, detail, ((Microsoft.Crm.Sdk.Messages.RelationshipAuditDetail)detail).TargetRecords);
                    //        auditRoleList.Add(auditRole);
                    //    }
                    //}
                }
                using (ExcelPackage excel = new ExcelPackage())
                {
                    excel.Workbook.Worksheets.Add("ApolloLastLogin");
                    var headerRow = new List<string[]>()
                      {
                        new string[] {  "User", "Last Login" }
                      };


                    // Target a worksheet
                    var worksheet = excel.Workbook.Worksheets["ApolloLastLogin"];
                    //string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";
                    //worksheet.Cells[headerRange].LoadFromArrays(headerRow);
                    worksheet.Cells[1, 1].LoadFromCollection(userList, true);
                    FileInfo excelFile = new FileInfo(@"Y:\Ketan\Role\Apollo.xlsx");
                    excel.SaveAs(excelFile);
                }
            }
        }
        public static AuditRole DisplayAuditDetails(CrmServiceClient service, AuditDetail detail,Microsoft.Xrm.Sdk.DataCollection<EntityReference> Target)
        {
            AuditRole auditroleObject = new AuditRole();
            var record = detail.AuditRecord;
            auditroleObject.AuditId = record.Id;

            auditroleObject.AssignedOn = (DateTime)record.Attributes["createdon"];
            auditroleObject.Role = string.Join(",", Target.Select(x => x.Name));
            auditroleObject.AsignedBy=((EntityReference)record.Attributes["userid"]).Name;
            string ObjectName = GetFullNameFromId(service, ((EntityReference)record.Attributes["objectid"]).Id);
            auditroleObject.User = ObjectName;
            Console.WriteLine("User :   {0} \n Role :   {1}", ObjectName, auditroleObject.Role);
            return auditroleObject;


        }
        public static string GetFullNameFromId(CrmServiceClient service, Guid userId)
        {
            
            Entity user= service.Retrieve("systemuser",userId,new ColumnSet(new string[]{"fullname"}));
            string fullname = user.GetAttributeValue<string>("fullname");
            return fullname;

        }
            public static EntityCollection GetUsers(CrmServiceClient service)
        {
            Console.WriteLine("IN: GetUsers()");
            string fetchXML = @"<fetch>
  <entity name='systemuser' >
    <attribute name='fullname' />
    <attribute name='systemuserid' />
    <filter>
      <condition attribute='isdisabled' operator='eq' value='0' />
    </filter>
  </entity>
</fetch>";
            EntityCollection users = service.RetrieveMultiple(new FetchExpression(fetchXML));
            Console.WriteLine("{0} Enabled Users Retrieved", users.Entities.Count);
            Console.WriteLine("OUT: GetUsers()");
            return users;
        }
        public static CrmServiceClient Connect(string name)
        {
            CrmServiceClient service = null;
            // Try to create via connection string. 
            service = new CrmServiceClient(GetConnectionStringFromAppConfig("Connect"));
            return service;
        }
        private static string GetConnectionStringFromAppConfig(string name)
        {
            //Verify cds/App.config contains a valid connection string with the name.
            try
            {
                return ConfigurationManager.ConnectionStrings[name].ConnectionString;
            }
            catch (Exception)
            {
                Console.WriteLine("You can set connection data in cds/App.config before running this sample. - Switching to Interactive Mode");
                return string.Empty;
            }
        }
    }
}
