using System;
using System.Configuration;
using System.Web.Http;
using TnOutlookWebApp.Models;
using Newtonsoft.Json;
using System.Diagnostics;

namespace TnOutlookWebApp.Controllers
{
    public class TasksController : ApiController
    {
        ExchangeHelper exchangeHelper;
        CrmHelper crmHelper;

        string userMail = ConfigurationManager.AppSettings["exchangeServiceUsername"];
        string password = ConfigurationManager.AppSettings["exchangeServicePass"];
        string crmUri = ConfigurationManager.AppSettings["organizationServiceUri"];
        string crmUser = ConfigurationManager.AppSettings["organizationServiceUsername"];
        string crmPass = ConfigurationManager.AppSettings["organizationServicePass"];

        [HttpPost]
        public string CreateTaskInOutlook(TaskEntity taskEntity)
        {
            Trace.TraceInformation("Create task in outlook");
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            string newOwnerMail = crmHelper.GetUserMailByGuid(taskEntity.NewTaskOwnerId);
            taskEntity.OutlookId = exchangeHelper.CreateNewOutlookTask(taskEntity, newOwnerMail);
            crmHelper.UpdateCrmTask(taskEntity);
            return JsonConvert.SerializeObject(taskEntity.OutlookId);
        }

        [HttpPost]
        public string UpdateTaskInOutlook(TaskEntity taskEntity)
        {
            Trace.TraceInformation("Update task in outlook");

            if (!InitializeHelpers())
                return "Initialize helpers fail";
            string outlookTaskId = crmHelper.GetOutlookId(taskEntity);            
            if (taskEntity.NewTaskOwnerId == Guid.Empty)
            {
                taskEntity.OutlookId = outlookTaskId;
                exchangeHelper.UpdateOutlookTask(taskEntity, outlookTaskId);
                crmHelper.UpdateCrmTask(taskEntity);
            }
            else
            {
                string newOwnerMail = crmHelper.GetUserMailByGuid(taskEntity.NewTaskOwnerId);
                string newOutlookTaskId = exchangeHelper.CreateNewOutlookTask(taskEntity, newOwnerMail);
                string oldOutlookIdForDelete = outlookTaskId;
                taskEntity.OutlookId = newOutlookTaskId;
                crmHelper.UpdateCrmTask(taskEntity);
                exchangeHelper.DeleteOutlookTaskById(oldOutlookIdForDelete);
            }
            return "Update success";
        }

        [HttpPost]
        public string UpdateTaskInCrm(Outlook outlook)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            TaskEntity outlookTask = exchangeHelper.GetTaskFromOutlook(outlook.outlookId);
            outlookTask.CrmId = crmHelper.GetCrmIdByOutlookId(outlookTask);
            string result = crmHelper.UpdateCrmTask(outlookTask);
            return result;
        }

        [HttpPost]
        public string DeleteTaskInOutlook(TaskEntity task)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            var res = exchangeHelper.DeleteOutlookTaskById(task.OutlookId);
            return JsonConvert.SerializeObject(new { result = res });
        }

        private bool InitializeHelpers()
        {
            if (exchangeHelper == null)
                exchangeHelper = new ExchangeHelper(userMail, password);
            if (crmHelper == null)
                crmHelper = new CrmHelper(crmUser, crmPass, crmUri);

            if (exchangeHelper != null && crmHelper != null)
                return true;
            else
                return false;
        }
    }
}