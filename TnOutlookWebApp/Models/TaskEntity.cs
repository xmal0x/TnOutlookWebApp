using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TnOutlookWebApp.Models
{
    public class TaskEntity
    {
        public Guid CrmId { get; set; }
        public string OutlookId {get;set;}
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime DuoDate { get; set; }
        public Guid OwnerId { get; set; }
        public Guid NewTaskOwnerId { get; set; }
        public TaskStatusCode TaskStatus { get; set; }
    }
}