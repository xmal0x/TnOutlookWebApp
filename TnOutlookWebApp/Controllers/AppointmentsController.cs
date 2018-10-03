using System;
using System.Configuration;
using System.Web.Http;
using TnOutlookWebApp.Models;
using Newtonsoft.Json;

namespace TnOutlookWebApp.Controllers
{
    public class AppointmentsController : ApiController
    {
        ExchangeHelper exchangeHelper;
        CrmHelper crmHelper;

        string userMail = ConfigurationManager.AppSettings["exchangeServiceUsername"];
        string password = ConfigurationManager.AppSettings["exchangeServicePass"];
        string crmUri = ConfigurationManager.AppSettings["organizationServiceUri"];
        string crmUser = ConfigurationManager.AppSettings["organizationServiceUsername"];
        string crmPass = ConfigurationManager.AppSettings["organizationServicePass"];

        [HttpPost]
        public string CreateAppointmentInOutlook(AppointmentEntity appointmentEntity)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            appointmentEntity.OutlookId = exchangeHelper.CreateNewOutlookAppointment(appointmentEntity);
            crmHelper.UpdateCrmAppointment(appointmentEntity);
            return JsonConvert.SerializeObject(appointmentEntity.OutlookId);
        }

        [HttpPost]
        public string UpdateAppointmentInOutlook(AppointmentEntity appointmentEntity)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            appointmentEntity.OutlookId = crmHelper.GetOutlookId(appointmentEntity);
            exchangeHelper.UpdateOutlookAppointment(appointmentEntity);
            crmHelper.UpdateCrmAppointment(appointmentEntity);
            return "Update appointment success";
        }

        [HttpPost]
        public string UpdateAppointmentInCrm(string outlookId)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            AppointmentEntity outlookAppointment = exchangeHelper.GetAppointmentFromOutlook(outlookId);
            string result = crmHelper.UpdateCrmAppointment(outlookAppointment);
            return result;
        }
        private bool InitializeHelpers()
        {
            if (exchangeHelper == null)
                exchangeHelper = new ExchangeHelper(userMail, password);
            if (crmHelper == null)
                crmHelper = new CrmHelper(crmUser, crmPass, crmUri);
            if (exchangeHelper != null && crmHelper != null)
                return true;
            else
                return false;
        }



    }
}