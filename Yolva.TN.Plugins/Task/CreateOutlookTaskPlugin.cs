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
    public class CreateOutlookTaskPlugin : PluginBase
    {
        private string serviceUrl = "http://azuretaskappweb.azurewebsites.net/api/tasks/CreateTaskInOutlook";
        public CreateOutlookTaskPlugin() : base("", "")
        {
        }
        public CreateOutlookTaskPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }
        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var task = pluginContext.TargetImageEntity;
            if (task == null)
                return;

            Guid crmId = Guid.Empty;
            string outlookId = string.Empty;
            string subject = string.Empty;
            string body = string.Empty;
            DateTime duoDate = DateTime.Now;
            Guid ownerId = Guid.Empty;
            Guid newTaskOwnerId = Guid.Empty;


            crmId = task.Id;
            ownerId = pluginContext.UserId;
            if (task.Contains("subject"))
            {
                subject = task["subject"].ToString();
            }
            if (task.Contains("description"))
            {
                body = task["description"].ToString();
            }
            if (task.Contains("scheduledend"))
            {
                duoDate = DateTime.Parse(task["scheduledend"].ToString());
            }
            if (task.Contains("ownerid"))
            {
                newTaskOwnerId = ((EntityReference)task["ownerid"]).Id;
            }

            CreateTaskInOutlook(new TaskEntity { CrmId = crmId, OutlookId = outlookId, Subject = subject, Body = body, DuoDate = duoDate, OwnerId = ownerId, NewTaskOwnerId = newTaskOwnerId }, serviceUrl);

        }
        public static string CreateTaskInOutlook(TaskEntity task, string serviceUrl)
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
    }
}
