using System;
using System.Configuration;
using System.Web.Http;
using TnOutlookWebApp.Models;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Diagnostics;

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
            Trace.TraceInformation("CreateAppointmentInOutlook");
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            if (!string.IsNullOrEmpty(appointmentEntity.OutlookId))
            {
                Trace.TraceInformation("CreateAppointmentInOutlook: appointment exist in outlook");
                return appointmentEntity.OutlookId;
            }
            appointmentEntity.OutlookId = exchangeHelper.CreateNewOutlookAppointment(appointmentEntity);
            //crmHelper.UpdateCrmAppointment(appointmentEntity);
            return JsonConvert.SerializeObject(appointmentEntity.OutlookId);
        }

        [HttpPost]
        public string UpdateAppointmentInOutlook(AppointmentEntity appointmentEntity)
        {
            Trace.TraceInformation("UpdateAppointmentInOutlook");

            if (!InitializeHelpers())
                return "Initialize helpers fail";
            appointmentEntity.OutlookId = crmHelper.GetOutlookId(appointmentEntity);
            exchangeHelper.UpdateOutlookAppointment(appointmentEntity);
            //crmHelper.UpdateCrmAppointment(appointmentEntity);
            return "Update appointment success";
        }

        [HttpPost]
        public string CreateAppointmentInCrm(Outlook outlookId)
        {
            Trace.TraceInformation("CreateAppointmentInCrm");

            if (!InitializeHelpers())
                return "Initialize helpers fail";
            if(crmHelper.GetCrmIdByOutlookId(new AppointmentEntity { OutlookId = outlookId.outlookId }) != Guid.Empty)
            {
                Trace.TraceInformation("CreateAppointmentInCrm: appointmtnt exist");

                return "Appointment is exist";
            }
            AppointmentEntity outlookAppointment = exchangeHelper.GetAppointmentFromOutlook(outlookId.outlookId);
            Guid result = crmHelper.CreateCrmAppointment(outlookAppointment);
            return result.ToString();
        }

        [HttpPost]
        public string UpdateAppointmentInCrm(Outlook outlookId)
        {
            Trace.TraceInformation("UpdateAppointmentInCrm");

            if (!InitializeHelpers())
                return "Initialize helpers fail";
            List<AttendeesResponse> responses = exchangeHelper.GetResponseStatus(outlookId.outlookId);
            var appointmentEntity = exchangeHelper.GetAppointmentFromOutlook(outlookId.outlookId);
            appointmentEntity.CrmId = crmHelper.GetCrmIdByOutlookId(appointmentEntity);
            crmHelper.UpdateResponsesInCrm(appointmentEntity, responses);
            return "UpdateAppointmentInCrm success";
        }

        [HttpPost]
        public string IsCrmAppointmentExist(Outlook outlook)
        {
            if (!InitializeHelpers())
                return "Initialize helpers fail";
            var appExist = new ExistResponse { IsExist = crmHelper.GetCrmIdByOutlookId(new AppointmentEntity { OutlookId = outlook.outlookId }) != Guid.Empty };
            return JsonConvert.SerializeObject(appExist);
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