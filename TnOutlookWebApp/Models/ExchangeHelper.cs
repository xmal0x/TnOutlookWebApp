using Microsoft.Exchange.WebServices.Data;
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
            return outlookTask.Id.UniqueId;
        }

        internal string CreateNewOutlookAppointment(AppointmentEntity appointmentEntity)
        {
            Appointment appointment = new Appointment(exchangeService)
            {
                Start = appointmentEntity.Start,
                End = appointmentEntity.End,
                Subject = appointmentEntity.Subject,
                Body = new MessageBody(BodyType.Text, appointmentEntity.Body),
                Location = appointmentEntity.Location,                
            };
            foreach (var attendees in appointmentEntity.RequiredAttendeesEmails)
                appointment.RequiredAttendees.Add(attendees);
            appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);            
            return appointment.Id.UniqueId;
        }

        internal string UpdateOutlookAppointment(AppointmentEntity appointmentEntity)
        {
            Appointment outlookAppointment = Appointment.Bind(exchangeService, appointmentEntity.OutlookId);
            if (!string.IsNullOrEmpty(appointmentEntity.Location))
                outlookAppointment.Location = appointmentEntity.Location;
            if (!string.IsNullOrEmpty(appointmentEntity.Subject))
                outlookAppointment.Subject = appointmentEntity.Subject;
            if (!string.IsNullOrEmpty(appointmentEntity.Body))
                outlookAppointment.Body = appointmentEntity.Body;
            if (appointmentEntity.Start != null)
                outlookAppointment.Start = appointmentEntity.Start;
            if (appointmentEntity.End != null)
                outlookAppointment.End = appointmentEntity.End;
            outlookAppointment.Update(ConflictResolutionMode.AlwaysOverwrite);
            return "Appointment update success";
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

        internal AppointmentEntity GetAppointmentFromOutlook(string outlookId)
        {
            throw new NotImplementedException();
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