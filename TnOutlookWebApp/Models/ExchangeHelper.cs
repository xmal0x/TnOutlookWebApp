﻿using Microsoft.Exchange.WebServices.Data;
using System;
using System.Web.Helpers;
using System.Web.Mvc;

namespace TnOutlookWebApp.Models
{
    public class ExchangeHelper
    {
        private ExchangeService exchangeService;
        private string userMail = string.Empty;
        private string password = string.Empty;
        public ExchangeHelper(string userMail, string password)
        {
            this.userMail = userMail;
            this.password = password;
            exchangeService = new ExchangeService(ExchangeVersion.Exchange2013_SP1, TimeZoneInfo.Utc);
            exchangeService.Credentials = new WebCredentials(this.userMail, this.password);
            exchangeService.TraceEnabled = true;
            exchangeService.AutodiscoverUrl(userMail, RedirectionUrlValidationCallback);
        }

        public string CreateNewOutlookTask(TaskEntity taskEntity, string ownerUserMail)
        {
            Task outlookTask = new Task(exchangeService)
            {
                Subject = taskEntity.Subject,
                Body = taskEntity.Body,
                DueDate = taskEntity.DuoDate                
            };
            outlookTask.Save(new FolderId(WellKnownFolderName.Tasks, ownerUserMail));
            return outlookTask.Id.ToString();
        }

        internal string UpdateOutlookTask(TaskEntity taskEntity, string outlookTaskId)
        {
            Task outlookTask = Task.Bind(exchangeService, new ItemId(outlookTaskId));
            if (!string.IsNullOrEmpty(taskEntity.Subject))
                outlookTask.Subject = taskEntity.Subject;
            if (!string.IsNullOrEmpty(taskEntity.Body))
                outlookTask.Body = taskEntity.Body;
            if (taskEntity.DuoDate != null)
                outlookTask.DueDate = taskEntity.DuoDate;
            outlookTask.Update(ConflictResolutionMode.AlwaysOverwrite);
            return "Update success";
        }

        internal void DeleteOutlookTaskById(string oldOutlookIdForDelete)
        {
            Task task = Task.Bind(exchangeService, new ItemId(oldOutlookIdForDelete));
            task.Delete(DeleteMode.SoftDelete);
        }

        internal TaskEntity GetTaskFromOutlook(string outlookId)
        {
            Task outlookTask = Task.Bind(exchangeService, new ItemId(outlookId));
            
            if (outlookTask == null)
                return null;
            TaskEntity task = new TaskEntity()
            {
                Subject = outlookTask.Subject,
                Body = outlookTask.Body.Text,
                DuoDate = outlookTask.DueDate == null ? DateTime.Now : (DateTime)outlookTask.DueDate,
                OutlookId =  outlookId
            };
            var currentStatus = outlookTask.Status;
            switch (currentStatus)
            {
                case TaskStatus.NotStarted:
                    task.TaskStatus = TaskStatusCode.Open;
                    break;
                case TaskStatus.InProgress:
                    task.TaskStatus = TaskStatusCode.Open;
                    break;
                case TaskStatus.Completed:
                    task.TaskStatus = TaskStatusCode.Complited;
                    break;
                case TaskStatus.WaitingOnOthers:
                    task.TaskStatus = TaskStatusCode.Open;
                    break;
                case TaskStatus.Deferred:
                    task.TaskStatus = TaskStatusCode.Open;
                    break;
                default:
                    task.TaskStatus = TaskStatusCode.Open;
                    break;
            }
            return task;
        }

        private bool RedirectionUrlValidationCallback(String redirectionUrl)
        {
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;

        }


    }
}