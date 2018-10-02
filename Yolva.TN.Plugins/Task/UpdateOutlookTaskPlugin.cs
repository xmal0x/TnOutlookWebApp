using Microsoft.Xrm.Sdk;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;
using Yolva.TN.Plugins.Common;

namespace Yolva.TN.Plugins.Task
{
    public class UpdateOutlookTaskPlugin : PluginBase
    {
        private string serviceUrl = "http://azuretaskappweb.azurewebsites.net/api/tasks/UpdateTaskInOutlook";
        public UpdateOutlookTaskPlugin() : base("", "")
        {
        }
        public UpdateOutlookTaskPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var task = pluginContext.TargetImageEntity;
            if (task == null)
                return;

            TaskEntity updateTask = new TaskEntity();
            Entity oldTask = pluginContext.OrganizationService.Retrieve("task", task.Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(true));
            if (oldTask == null)
            {
                throw new InvalidPluginExecutionException("Error then try get old task data");
            }

            updateTask.CrmId = task.Id;

            updateTask.Subject = oldTask["subject"].ToString();
            updateTask.Body = oldTask["description"].ToString();
            updateTask.DuoDate = DateTime.Parse(oldTask["scheduledend"].ToString());
            updateTask.NewTaskOwnerId = Guid.Empty;

            //updateTask.OwnerId = pluginContext.UserId;
            if (task.Contains("subject"))
            {
                updateTask.Subject = task["subject"].ToString();

            }
            if (task.Contains("description"))
            {
                updateTask.Body = task["description"].ToString();
            }
            if (task.Contains("scheduledend"))
            {
                updateTask.DuoDate = DateTime.Parse(task["scheduledend"].ToString());
            }
            if (task.Contains("ownerid"))
            {
                updateTask.NewTaskOwnerId = ((EntityReference)task["ownerid"]).Id;
            }
            UpdateTaskInOutlook(updateTask, serviceUrl);

        }

        public static string UpdateTaskInOutlook(TaskEntity task, string serviceUrl)
        {
            var ApiServiceUrl = serviceUrl;
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.ContentType] = "application/json";
                var jsonObj = JsonConvert.SerializeObject(task);
                var dataString = client.UploadString(ApiServiceUrl, jsonObj);
                //var data = JsonConvert.DeserializeObject<TaskEntity>(dataString);
                return "Success";
            }
        }

        private async void UpdateTaskInOutlook(string crmTaskId, string taskDescription, string taskSubject, string ownerId, string dateEnd)
        {
            var client = new HttpClient();
            string url = "http://azuretaskwebapp.azurewebsites.net/api/task/updatetask/" + crmTaskId + ";" + taskDescription + ";" + taskSubject + ";" + ownerId + ";" + dateEnd + ";";

            var response = await client.GetAsync(url);
            var content = await response.Content.ReadAsStringAsync();
        }
    }
}
