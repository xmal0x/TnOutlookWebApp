
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Yolva.TN.Plugins.Common
{
    public class AppointmentEntity
    {        
        public string OutlookId { get; set; }
        public string CrmId { get; set; }
        public List<string> RequiredAttendeesEmails { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string Location { get; set; }

        public AppointmentEntity()
        {
            Body = string.Empty;
            CrmId = string.Empty;
            Start = DateTime.Now.AddHours(1);
            End = DateTime.Now.AddHours(2);
            Location = string.Empty;
            OutlookId = string.Empty;
            RequiredAttendeesEmails = null;
            Subject = string.Empty;
        }
    }
}