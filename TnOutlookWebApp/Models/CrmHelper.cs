﻿using Microsoft.Xrm.Sdk;
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
            return GetFieldValueFromEntity("systemuser", "internalemailaddress", guid);
        }

        private string GetFieldValueFromEntity(string entityLogicalName, string fieldName, Guid guid)
        {
            string value = null;
            var entity = organizationService.Retrieve(entityLogicalName, guid, new ColumnSet(fieldName));
            if (entity != null)
                value = entity[fieldName].ToString();
            return value;
        }

        internal string UpdateCrmAppointment(AppointmentEntity newAppointment)
        {
            try
            {
                Entity updAppointment = new Entity(newAppointment.CrmEntityLogicalName);
                updAppointment.Id = newAppointment.CrmId;
                if (!string.IsNullOrEmpty(newAppointment.Body))
                {
                    updAppointment["description"] = newAppointment.Body;
                }
                if (newAppointment.End != null)
                {
                    updAppointment["scheduledend"] = newAppointment.End;
                }
                if (newAppointment.Start != null)
                {
                    updAppointment["scheduledstart"] = newAppointment.Start;
                }
                if (!string.IsNullOrEmpty(newAppointment.Location))
                {
                    updAppointment["location"] = newAppointment.Location;
                }
                if (!string.IsNullOrEmpty(newAppointment.OutlookId))
                {
                    updAppointment["ylv_outlookid"] = newAppointment.OutlookId;
                }
                if (!string.IsNullOrEmpty(newAppointment.Subject))
                {
                    updAppointment["subject"] = newAppointment.Subject;
                }
                organizationService.Update(updAppointment);
                return "Appointment update success";
            }
            catch (Exception ex)
            {
                return "Appointment update fail\n" + ex.Message;
            }
        }

        internal Guid GetCrmIdByOutlookId(IntegrationEntity entity)
        {
            string fetch = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" + 
                              "<entity name='" + entity.CrmEntityLogicalName + "'>" +
                                "<order attribute='subject' descending='false' />" + 
                                "<filter type='and'>" + 
                                  "<condition attribute='ylv_outlookid' operator='eq' value='" + entity.OutlookId + "' />" + 
                                "</filter>" +
                              "</entity>" +
                            "</fetch>";
            var crmEntity = organizationService.RetrieveMultiple(new FetchExpression(fetch)).Entities.FirstOrDefault();
            return crmEntity == null ? Guid.Empty : crmEntity.Id;
        }

        internal string GetOutlookId(IntegrationEntity entity)
        {
            var crmEntity = organizationService.Retrieve(entity.CrmEntityLogicalName, entity.CrmId, new ColumnSet("ylv_outlookid"));
            return crmEntity["ylv_outlookid"] == null ? null : crmEntity["ylv_outlookid"].ToString();
        }

        internal string UpdateCrmTask(TaskEntity newTask)
        {
            try
            {
                Entity updCrmTask = new Entity(newTask.CrmEntityLogicalName);
                updCrmTask.Id = newTask.CrmId;
                if(newTask.TaskStatus != null)
                {
                    updCrmTask["statecode"] = new OptionSetValue((int)newTask.TaskStatus);
                }
                if (!string.IsNullOrEmpty(newTask.Subject))
                {
                    updCrmTask["subject"] = newTask.Subject;
                }
                if (!string.IsNullOrEmpty(newTask.Subject))
                {
                    updCrmTask["description"] = newTask.Body;
                }
                if(newTask.DuoDate != null)
                {
                    updCrmTask["scheduledend"] = newTask.DuoDate;
                }
                if (!string.IsNullOrEmpty(newTask.OutlookId))
                {
                    updCrmTask["ylv_outlookid"] = newTask.OutlookId;
                }
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