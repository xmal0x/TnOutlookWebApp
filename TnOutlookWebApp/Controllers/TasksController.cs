using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;
using System.Web.Http;
//using System.Web.Mvc;
using TnOutlookWebApp.Models;

namespace TnOutlookWebApp.Controllers
{
    public class TasksController : ApiController
    {
        ExchangeHelper exchangeHelper;
        CrmHelper crmHelper;
        AzureHelper azureHelper;
        string userMail = ConfigurationManager.AppSettings["exchangeServiceUsername"];
        string password = ConfigurationManager.AppSettings["exchangeServicePass"];
        string crmUri = ConfigurationManager.AppSettings["organizationServiceUri"];
        string crmUser = ConfigurationManager.AppSettings["organizationServiceUsername"];
        string crmPass = ConfigurationManager.AppSettings["organizationServicePass"];
        string azureConnectionString = ConfigurationManager.AppSettings["cloudStorageConnectionString"];

        string tasksTableName = "tasksTable";

        [HttpPost]
        public string CreateTaskInOutlook(TaskEntity taskEntity)
        {
            if (!InitializeHelpers())
                return "Error";
            string newOwnerMail = crmHelper.GetUserMailByGuid(taskEntity.NewTaskOwnerId);
            taskEntity.OutlookId = exchangeHelper.CreateNewOutlookTask(taskEntity, newOwnerMail);

            azureHelper.CreateTaskRecord(taskEntity, tasksTableName);
            return taskEntity.OutlookId;
        }

        [HttpPost]
        public string UpdateTaskInOutlook(TaskEntity taskEntity)
        {
            if (!InitializeHelpers())
                return "Error";
            string outlookTaskId = azureHelper.GetOutlookTaskId(taskEntity, tasksTableName);
            if(taskEntity.NewTaskOwnerId == Guid.Empty)
            {
                taskEntity.OutlookId = outlookTaskId;
                exchangeHelper.UpdateOutlookTask(taskEntity, outlookTaskId);
                azureHelper.UpdateTaskRecord(taskEntity, tasksTableName);
            }
            else
            {
                string newOwnerMail = crmHelper.GetUserMailByGuid(taskEntity.NewTaskOwnerId);
                
                string newOutlookTaskId = exchangeHelper.CreateNewOutlookTask(taskEntity, newOwnerMail);
                string oldOutlookIdForDelete = outlookTaskId;
                taskEntity.OutlookId = newOutlookTaskId;
                azureHelper.UpdateTaskRecord(taskEntity, tasksTableName);
                exchangeHelper.DeleteOutlookTaskById(oldOutlookIdForDelete);
            }
            return "Update success";            
        }

        [HttpPost]
        public string UpdateTaskInCrm(string outlookId)
        {
            //Task task = Task.Bind();

            if (!InitializeHelpers())
                return "Error";
            TaskEntity outlookTask = exchangeHelper.GetTaskFromOutlook(outlookId);
            outlookTask.CrmId = azureHelper.GetCrmTaskIdByOutlookId(outlookTask, tasksTableName);
            string result = crmHelper.UpdateCrmTask(outlookTask);
            azureHelper.UpdateTaskRecord(outlookTask, tasksTableName);
            return "result";
        }

        private bool InitializeHelpers()
        {
            if (exchangeHelper == null)
                exchangeHelper = new ExchangeHelper(userMail, password);
            if (crmHelper == null)
                crmHelper = new CrmHelper(crmUser, crmPass, crmUri);
            if (azureHelper == null)
                azureHelper = new AzureHelper(azureConnectionString);

            if (exchangeHelper != null && crmHelper != null && azureHelper != null)
                return true;
            else
                return false;
        }
    }
}