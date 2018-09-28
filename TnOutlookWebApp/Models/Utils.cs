using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace TnOutlookWebApp.Models
{
    public static class Utils
    {
        static string userMail = ConfigurationManager.AppSettings["exchangeServiceUsername"];
        static string password = ConfigurationManager.AppSettings["exchangeServicePass"];
        static string crmUri = ConfigurationManager.AppSettings["organizationServiceUri"];
        static string crmUser = ConfigurationManager.AppSettings["organizationServiceUsername"];
        static string crmPass = ConfigurationManager.AppSettings["organizationServicePass"];
        static string azureConnectionString = ConfigurationManager.AppSettings["cloudStorageConnectionString"];

        public static bool InitializeHelpers(ExchangeHelper exchangeHelper, CrmHelper crmHelper, AzureHelper azureHelper)
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