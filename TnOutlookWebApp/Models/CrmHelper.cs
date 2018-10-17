using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text.RegularExpressions;
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

        internal bool IsTaskNeedUpdate(TaskEntity outlookTask)
        {
            QueryExpression queryExpression = new QueryExpression()
            {
                EntityName = outlookTask.CrmEntityLogicalName,
                Criteria =
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression("subject", ConditionOperator.Equal, outlookTask.Subject),
                        new ConditionExpression("scheduledend", ConditionOperator.On, outlookTask.DuoDate),
                        new ConditionExpression("description", ConditionOperator.Equal, GetStringWithoutTags(outlookTask.Body)),
                        new ConditionExpression("ylv_outlookid", ConditionOperator.Equal, outlookTask.OutlookId),
                        new ConditionExpression("statecode", ConditionOperator.Equal, (int)outlookTask.TaskStatus)
                    }
                }
            };

            var task = organizationService.RetrieveMultiple(queryExpression).Entities.FirstOrDefault();
            return task == null ? true : false;
            
        }

        internal void UpdateResponsesInCrm(AppointmentEntity appointmentEntity, List<AttendeesResponse> outlookResponses)
        {
            ConditionExpression condition = new ConditionExpression();
            condition.AttributeName = "ylv_responses";
            condition.Operator = ConditionOperator.Equal;
            condition.Values.Add(appointmentEntity.CrmId.ToString());
            ColumnSet columns = new ColumnSet("ylv_contact", "ylv_responsetype");
            QueryExpression query1 = new QueryExpression();
            query1.ColumnSet = columns;
            query1.EntityName = "ylv_response";
            query1.Criteria.AddCondition(condition);

            EntityCollection associatedResponses = organizationService.RetrieveMultiple(query1);
            foreach(var response in associatedResponses.Entities)
            {
                if(response["ylv_contact"] != null)
                {
                    var contactResponseEntityRef = (EntityReference)response["ylv_contact"];
                    var contactResponseEntity = organizationService.Retrieve(contactResponseEntityRef.LogicalName, contactResponseEntityRef.Id, new ColumnSet("emailaddress1"));
                    foreach(var outlookResp in outlookResponses)
                    {
                        if(contactResponseEntity["emailaddress1"].ToString().ToUpper() == outlookResp.Email.ToUpper())
                        {
                            switch (outlookResp.Response)
                            {
                                case ResponseType.Accept:
                                    response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Accept);

                                    break;
                                case ResponseType.Discard:
                                    response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Discard);
                                    break;
                                case ResponseType.Unknown:
                                    response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Unknown);
                                    break;
                                default:
                                    response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Unknown);
                                    break;
                            }
                            organizationService.Update(response);
                        }
                    }
                    
                }
              

            }


        }

        internal Guid CreateCrmAppointment(AppointmentEntity outlookAppointment)
        {
            Entity crmAppointment = new Entity(outlookAppointment.CrmEntityLogicalName);
            crmAppointment["description"] = GetStringWithoutTags(outlookAppointment.Body);
            crmAppointment["scheduledend"] = outlookAppointment.End;
            crmAppointment["location"] = outlookAppointment.Location;
            crmAppointment["ylv_outlookid"] = outlookAppointment.OutlookId;
            crmAppointment["scheduledstart"] = outlookAppointment.Start;
            crmAppointment["subject"] = outlookAppointment.Subject;

            var crmAttendeesPartyCol = new EntityCollection();

            foreach (var email in outlookAppointment.RequiredAttendeesEmails)
            {
                var attendeesCrmEntity = GetCrmAttendeesByEmail(email);
                if(attendeesCrmEntity != null)
                {
                    var crmAttendeesList = new Entity("activityparty");
                    crmAttendeesList["partyid"] = new EntityReference(attendeesCrmEntity.LogicalName, attendeesCrmEntity.Id);
                    crmAttendeesPartyCol.Entities.Add(crmAttendeesList);                    
                }
            }
            
            crmAppointment["requiredattendees"] = crmAttendeesPartyCol;
            crmAppointment.Id = organizationService.Create(crmAppointment);

            foreach (var att in crmAttendeesPartyCol.Entities)
            {
                var response = new Entity("ylv_response");
                response["ylv_name"] = crmAppointment["subject"].ToString();
                response["ylv_contact"] = (EntityReference)att["partyid"];
                response["ylv_responsetype"] = new OptionSetValue((int)ResponseType.Unknown);
                response["ylv_responses"] = crmAppointment.ToEntityReference();
                response.Id = organizationService.Create(response);
            }

            return crmAppointment.Id;

        }

        private Entity GetCrmAttendeesByEmail(string email)
        {
            QueryExpression queryExpressionContact = new QueryExpression("contact");
            queryExpressionContact.Criteria.AddCondition("emailaddress1", ConditionOperator.Equal, email);
            queryExpressionContact.ColumnSet = new ColumnSet("emailaddress1");
            var contact = organizationService.RetrieveMultiple(queryExpressionContact).Entities.FirstOrDefault();
            if (contact != null)
                return contact;
            QueryExpression queryExpressionSysUser = new QueryExpression("systemuser");
            queryExpressionSysUser.Criteria.AddCondition("internalemailaddress", ConditionOperator.Equal, email);
            queryExpressionSysUser.ColumnSet = new ColumnSet("internalemailaddress");
            var sysUser = organizationService.RetrieveMultiple(queryExpressionSysUser).Entities.FirstOrDefault();
            //if (sysUser != null)
                return sysUser;

        }

        internal Guid GetCrmIdByOutlookId(IntegrationEntity entity)
        {
            QueryExpression queryExpression = new QueryExpression()
            {
                EntityName = entity.CrmEntityLogicalName,
                Criteria =
                {
                    FilterOperator = LogicalOperator.And,
                    Conditions =
                    {
                        new ConditionExpression("ylv_outlookid", ConditionOperator.Equal, entity.OutlookId)
                    }
                }
            };
            var crmEntity = organizationService.RetrieveMultiple(queryExpression).Entities.FirstOrDefault();
            return crmEntity == null ? Guid.Empty : crmEntity.Id;
        }

        internal string GetOutlookId(IntegrationEntity entity)
        {
            var crmEntity = organizationService.Retrieve(entity.CrmEntityLogicalName, entity.CrmId, new ColumnSet("ylv_outlookid"));
            return crmEntity["ylv_outlookid"] == null ? null : crmEntity["ylv_outlookid"].ToString();
        }

        internal string UpdateCrmTask(TaskEntity newTask)
        {
            Trace.TraceInformation("UpdateCrmTask");                     

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
                if (!string.IsNullOrEmpty(newTask.Body))
                {                    
                    updCrmTask["description"] = GetStringWithoutTags(newTask.Body);
                }
                if (newTask.DuoDate != null )
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



        private object GetStringWithoutTags(string body)
        {
            string bodyWithoutTags = Regex.Replace(body, "<[^>]*>", "\n");
            return bodyWithoutTags.Trim();
        }
    }
}