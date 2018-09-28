using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TnOutlookWebApp.Models
{
    public class AzureTaskEntity : TableEntity
    {
        public AzureTaskEntity()
        {
            this.PartitionKey = "partition";
            this.RowKey = Guid.NewGuid().ToString();
        }
        public string CrmId { get; set; }
        public string OutlookId { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public string DuoDate { get; set; }
        public string OwnerId { get; set; }
        public string NewTaskOwnerId { get; set; }
    }
}