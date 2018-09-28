using System;
using System.Configuration;
using System.Web.Http;
using TnOutlookWebApp.Models;

namespace TnOutlookWebApp.Controllers
{
    public class AppointmentsController : ApiController
    {
        ExchangeHelper exchangeHelper;
        CrmHelper crmHelper;
        AzureHelper azureHelper;

        string userMail = ConfigurationManager.AppSettings["exchangeServiceUsername"];
        string password = ConfigurationManager.AppSettings["exchangeServicePass"];
        string crmUri = ConfigurationManager.AppSettings["organizationServiceUri"];
        string crmUser = ConfigurationManager.AppSettings["organizationServiceUsername"];
        string crmPass = ConfigurationManager.AppSettings["organizationServicePass"];
        string azureConnectionString = ConfigurationManager.AppSettings["cloudStorageConnectionString"];

        string tasksTableName = "tasksTable";

        [HttpPost]
        public string CreateAppointmentInOutlook(AppointmentEntity appointmentEntity)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            appointmentEntity.OutlookId = exchangeHelper.CreateNewOutlookAppointment(appointmentEntity);
            //azureHelper.CreateAppointmentRecord(appointmentEntity, tasksTableName);
            return appointmentEntity.OutlookId;
        }

        [HttpPost]
        public string UpdateAppointmentInOutlook(AppointmentEntity appointmentEntity)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            //string outlookAppointmentId = azureHelper.
            return "";
        }

        [HttpPost]
        public string UpdateAppointmentInCrm(string outlookId)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            AppointmentEntity outlookAppointment = exchangeHelper.GetAppointmentFromOutlook(outlookId);
            outlookAppointment.CrmId = azureHelper.GetCrmAppointmentIdByOutlookId(outlookAppointment, tasksTableName);
            string result = crmHelper.UpdateCrmAppointment(outlookAppointment);
            azureHelper.UpdateAppointmentRecord(outlookAppointment, tasksTableName);
            return "Success update";
        }
        private bool InitializeHelpers()
        {
            if (exchangeHelper == null)
                exchangeHelper = new ExchangeHelper(userMail, password);
            if (crmHelper == null)
                crmHelper = new CrmHelper(crmUser, crmPass, crmUri);
            if (azureHelper == null)
                azureHelper = new AzureHelper(azureConnectionString);

            if (exchangeHelper != null && crmHelper != null && azureHelper != null)
                return true;
            else
                return false;
        }



    }
}