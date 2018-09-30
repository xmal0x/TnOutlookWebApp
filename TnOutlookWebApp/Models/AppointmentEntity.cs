using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TnOutlookWebApp.Models
{
    public class AppointmentEntity :TableEntity
    {
        public AppointmentEntity()
        {
            this.PartitionKey = DateTime.Now.ToString("yyyyMM");
            this.RowKey = DateTime.Now.Ticks.ToString();
        }

        public string OutlookId { get; set; }
        public string CrmId { get; set; }
        public List<string> RequiredAttendeesEmails { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Location { get; set; }
    }
}