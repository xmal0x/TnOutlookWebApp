using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Yolva.TN.Plugins.BaseClasses;
using Yolva.TN.Plugins.Common;

namespace Yolva.TN.Plugins.Appointment
{
    public class CreateOutlookAppointmentPlugin : PluginBase
    {
        private string serviceUrl = "http://azuretaskappweb.azurewebsites.net/api/Appointments/CreateAppointmentInOutlook";
        public CreateOutlookAppointmentPlugin() : base("", "")
        {
        }
        public CreateOutlookAppointmentPlugin(string unsecure = "", string secure = "") : base(unsecure, secure)
        {
        }

        protected override void ExecuteBusinessLogic(PluginContext pluginContext)
        {
            var appointment = pluginContext.TargetImageEntity;
            if (appointment == null)
                return;

            AppointmentEntity newAppointment = new AppointmentEntity();

            newAppointment.CrmId = appointment.Id.ToString();
            if (appointment.Contains("subject"))
                newAppointment.Subject = appointment["subject"].ToString();
            if (appointment.Contains("location"))
                newAppointment.Location = appointment["location"].ToString();
            if (appointment.Contains("description"))
                newAppointment.Body = appointment["description"].ToString();
            if (appointment.Contains("scheduledstart"))
                newAppointment.Start = DateTime.Parse(appointment["scheduledstart"].ToString());
            if (appointment.Contains("scheduledend"))
                newAppointment.End = DateTime.Parse(appointment["scheduledend"].ToString());
            if (appointment.Contains("ylv_emailness"))
            {
                foreach (var attendees in appointment["ylv_emailness"].ToString().Split(new string[] { ";" }, StringSplitOptions.RemoveEmptyEntries))
                    newAppointment.RequiredAttendeesEmails.Add(attendees.Trim());
            }

            CreateAppointmentInOutlook(newAppointment, serviceUrl);
        }
        public static string CreateAppointmentInOutlook(AppointmentEntity appointment, string serviceUrl)
        {
            var ApiServiceUrl = serviceUrl;
            using (WebClient client = new WebClient())
            {
                client.Headers[HttpRequestHeader.ContentType] = "application/json";
                var jsonObj = JsonConvert.SerializeObject(appointment);
                var dataString = client.UploadString(ApiServiceUrl, jsonObj);
                return "Success";
            }
        }
    }
}
