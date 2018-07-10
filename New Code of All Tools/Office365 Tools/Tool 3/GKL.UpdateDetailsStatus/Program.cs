using System;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Net;
using log4net;

namespace GKL.UpdateDetailsStatus
{
    class Program
    {
        static string coversheetHeaderList = string.Empty;
        static string coversheetDetailsList = string.Empty;
        static string siteurl = string.Empty;
        static string username = string.Empty;
        static string password = string.Empty;
        static string domain = string.Empty;
        static bool isCloud;
        static ClientContext context;

        static void Main(string[] args)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Started Processing..");
                Logger.log.Info("=========== Started Processing of Details Updation ====================");
                ReadConfigurationValues();
                Logger.log.Info("=========== Completed Processing of Details Updation ====================");
                Console.WriteLine("Processing Completed, Please close the window..!");
              //  Console.ReadKey();
            }
            catch(Exception ex)
            {
                Logger.log.Error("Error in updating the details " + ex.Message);
            }
        }

        private static void ReadConfigurationValues()
        {
            Logger.log.Info("Reading Configuration values from config");
            coversheetHeaderList = ConfigurationManager.AppSettings.Get("CoversheetHeaderList");
            coversheetDetailsList = ConfigurationManager.AppSettings.Get("CoversheetDetailsList");
            siteurl = ConfigurationManager.AppSettings.Get("SiteUrl");
            domain = ConfigurationManager.AppSettings.Get("Domain");
            username = ConfigurationManager.AppSettings.Get("UserName");
            password = ConfigurationManager.AppSettings.Get("Password");
            isCloud = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("CloudEnvironment"));
            if (!string.IsNullOrEmpty(siteurl))
            {
                context = new ClientContext(siteurl);
            }
            else
            {
                Logger.log.Info("Site Url set to null in the configuration");
            }

            if (isCloud)
            {
                Logger.log.Info("Environment configuration set for Cloud");
                System.Security.SecureString pwdSecureString = new System.Security.SecureString();
                if (!string.IsNullOrEmpty(password))
                {
                    foreach (char c in password)
                    {
                        pwdSecureString.AppendChar(c);
                    }

                    context.Credentials = new SharePointOnlineCredentials(username, pwdSecureString);
                }
            }
            else
            {
                Logger.log.Info("Environment configuration set for On premise");
                context.Credentials = new NetworkCredential(username, password, domain);
            }

            Web web = context.Web;
            context.Load(web, w => w.Title, w => w.Url);
            List listHeader = web.Lists.GetByTitle(coversheetHeaderList);
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                        "<Value Type='Number'>0</Value></Geq></Where></Query></View>";
            ListItemCollection collection = listHeader.GetItems(camlQuery);
            context.Load(collection, items => items.Include(
                     item => item.Id,
                     item => item["Title"],
                      item => item["Vendor_x0020_Name"],
                     item => item["Accounting_x0020_Approval"]
                     ));
            context.ExecuteQuery();
            Logger.log.Info("Context created for the web and header list");
            foreach (ListItem item in collection)
            {
                UpdateDetailsList(web, item.Id.ToString(), item["Accounting_x0020_Approval"].ToString());
            }          
        }

        private static void UpdateDetailsList(Web web, string sourceItemID, string paymentStatus)
        {
            try
            {
                Logger.log.Info("Updating process started for the header item " + sourceItemID);
                List listDetails = web.Lists.GetByTitle(coversheetDetailsList);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='SourceItemId'/>" +
                            "<Value Type='Text'>" + sourceItemID + "</Value></Eq></Where></Query></View>";
                ListItemCollection collection = listDetails.GetItems(camlQuery);
                context.Load(collection, items => items.Include(
                         item => item.Id,
                         item => item["Accounting_x0020_Approval"]
                         ));
                context.ExecuteQuery();
                if (collection.Count > 0)
                {
                    foreach (ListItem item in collection)
                    {
                        Logger.log.Info("Updating details for the item :" + sourceItemID);
                        item["Accounting_x0020_Approval"] = paymentStatus;
                        item.Update();
                        context.Load(item);
                    }
                    context.ExecuteQuery();
                    Logger.log.Info("Details Updation Completed ");
                }
                else
                {
                    Logger.log.Info("No details found with the source item " + sourceItemID);
                }
            }
            catch(Exception ex)
            {
                Logger.log.Error("Error in Updating Details List :" + ex.Message);

            }
        }
    }

    public static class Logger
    {
        public static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        static Logger()
        {
            log4net.Config.XmlConfigurator.Configure();
        }
    }
}
