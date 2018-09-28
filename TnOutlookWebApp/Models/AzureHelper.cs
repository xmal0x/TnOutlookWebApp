using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Newtonsoft.Json;

namespace TnOutlookWebApp.Models
{
    public class AzureHelper
    {
        private string cloudStorageConnectionString = string.Empty;
        CloudStorageAccount storageAccount;
        CloudTableClient tableClient;

        public AzureHelper(string connectionString)
        {
            cloudStorageConnectionString = connectionString;
            storageAccount = CloudStorageAccount.Parse(cloudStorageConnectionString);
            tableClient = storageAccount.CreateCloudTableClient();
        }

        public void CreateTaskRecord(TaskEntity taskEntity, string tableName)
        {
            var table = tableClient.GetTableReference(tableName);
            table.CreateIfNotExists();

            var task = new AzureTaskEntity()
            {
                CrmId = taskEntity.CrmId.ToString(),
                OutlookId = taskEntity.OutlookId,
                Subject = taskEntity.Subject,
                Body = taskEntity.Body,
                DuoDate = taskEntity.DuoDate.ToString(),
                OwnerId = taskEntity.OwnerId.ToString(),
                NewTaskOwnerId = taskEntity.NewTaskOwnerId.ToString()
            };

            var insertOperation = TableOperation.Insert(task);
            table.Execute(insertOperation);
        }

        internal string GetOutlookTaskId(TaskEntity taskEntity, string tableName)
        {
            var table = tableClient.GetTableReference(tableName);

            TableQuery<AzureTaskEntity> query = new TableQuery<AzureTaskEntity>().Where(TableQuery.GenerateFilterCondition("CrmId", QueryComparisons.Equal, taskEntity.CrmId.ToString()));

            var azureEntity = table.ExecuteQuery(query).FirstOrDefault();
            if (azureEntity != null)
            {
               return azureEntity.OutlookId;
            }
            return "Error";
        }

        internal void UpdateTaskRecord(TaskEntity updateTask, string tableName)
        {
            var table = tableClient.GetTableReference(tableName);

            TableQuery<AzureTaskEntity> query = new TableQuery<AzureTaskEntity>().Where(TableQuery.GenerateFilterCondition("CrmId", QueryComparisons.Equal, updateTask.CrmId.ToString()));

            var azureEntity = table.ExecuteQuery(query).FirstOrDefault();
            if (azureEntity != null)
            {
                if(!string.IsNullOrEmpty(updateTask.Subject))
                    azureEntity.Subject = updateTask.Subject;
                if (!string.IsNullOrEmpty(updateTask.Body))
                    azureEntity.Body = updateTask.Body;
                if (!string.IsNullOrEmpty(updateTask.DuoDate.ToString()))
                    azureEntity.DuoDate = updateTask.DuoDate.ToString();
                if (!string.IsNullOrEmpty(updateTask.OutlookId))
                    azureEntity.OutlookId = updateTask.OutlookId;
                if (updateTask.NewTaskOwnerId != Guid.Empty)
                    azureEntity.NewTaskOwnerId = updateTask.NewTaskOwnerId.ToString();
                if (updateTask.OwnerId != Guid.Empty)
                    azureEntity.OwnerId = updateTask.OwnerId.ToString();

                TableOperation updateOperation = TableOperation.Replace(azureEntity);
                table.Execute(updateOperation);
            }
        }

        internal Guid GetCrmTaskIdByOutlookId(TaskEntity taskEntity, string tableName)
        {
            var table = tableClient.GetTableReference(tableName);

            TableQuery<AzureTaskEntity> query = new TableQuery<AzureTaskEntity>().Where(TableQuery.GenerateFilterCondition("OutlookId", QueryComparisons.Equal, taskEntity.OutlookId.ToString()));

            var azureEntity = table.ExecuteQuery(query).FirstOrDefault();
            if (azureEntity != null)
            {
                return Guid.Parse(azureEntity.CrmId);
            }
            return Guid.Empty;
        }
    }
}