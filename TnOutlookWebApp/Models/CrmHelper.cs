using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Web;

namespace TnOutlookWebApp.Models
{
    public class CrmHelper
    {
        IOrganizationService organizationService;
        private string userName = string.Empty;
        private string password = string.Empty;
        private string orgUri = string.Empty;

        public CrmHelper(string userName, string pass, string orgUri)
        {
            this.userName = userName;
            this.password = pass;
            this.orgUri = orgUri;

            ClientCredentials clientCredentials = new ClientCredentials();
            clientCredentials.UserName.UserName = this.userName;
            clientCredentials.UserName.Password = this.password;

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            organizationService = (IOrganizationService)new OrganizationServiceProxy(new Uri(orgUri),
                null, clientCredentials, null);
        }

        public string GetUserMailByGuid(Guid guid)
        {
            return GetFieldValueFromEntity("systemuser", "regardingobjectid", guid);
            /*
            string email = string.Empty;
            var user = organizationService.Retrieve("systemuser", guid, new ColumnSet("internalemailaddress"));
            if (user != null)
                email = user["internalemailaddress"].ToString();
            return email;
            */
        }

        private string GetFieldValueFromEntity(string entityLogicalName, string fieldName, Guid guid)
        {
            string value = null;
            var entity = organizationService.Retrieve(entityLogicalName, guid, new ColumnSet(fieldName));
            if (entity != null)
                value = entity[fieldName].ToString();
            return value;
        }

        internal string UpdateCrmAppointment(AppointmentEntity outlookAppointment)
        {
            throw new NotImplementedException();
        }

        internal string UpdateCrmTask(TaskEntity outlookTask)
        {
            var crmTaskId = outlookTask.CrmId;
            var oldCrmTask = organizationService.Retrieve("task", crmTaskId, new ColumnSet(true));
            try
            {
                Entity updCrmTask = new Entity("task");
                updCrmTask.Id = oldCrmTask.Id;
                updCrmTask["statecode"] = new OptionSetValue((int)outlookTask.TaskStatus);
                updCrmTask["subject"] = outlookTask.Subject;
                updCrmTask["description"] = outlookTask.Body;
                updCrmTask["scheduledend"] = outlookTask.DuoDate;
                organizationService.Update(updCrmTask);
                return "Success update";
            }
            catch (Exception ex)
            {                
                return "Update fail\n" + ex.Message;
            }

        }
    }
}