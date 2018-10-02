using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TnOutlookWebApp.Models
{
    public abstract class IntegrationEntity
    {
        public string OutlookId { get; set; }
        public Guid CrmId { get; set; }
        public string CrmEntityLogicalName { get; set; }
    }
}