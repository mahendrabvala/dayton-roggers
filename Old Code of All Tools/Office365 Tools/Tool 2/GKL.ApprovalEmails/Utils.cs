using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Net.Mail;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;
using log4net;
namespace GKL.ApprovalEmails
{
    public static class Utils
    {
        public static void ReadConfigurationValues()
        {
            try
            {
                Logger.log.Info("Reading Configurations");
                Invoice._invoiceCoverSheetHeaderListName = ConfigurationManager.AppSettings.Get("CoversheetHeaderList");
                Invoice._orgListName = ConfigurationManager.AppSettings.Get("OrganizationList");
                Invoice._siteUrl = ConfigurationManager.AppSettings.Get("SiteUrl");
                Invoice._domain = ConfigurationManager.AppSettings.Get("Domain");
                Invoice._userName = ConfigurationManager.AppSettings.Get("UserName");
                Invoice._password = ConfigurationManager.AppSettings.Get("Password");
                Invoice._runforCloud = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("CloudEnvironment"));
                Invoice._displayFormLink = ConfigurationManager.AppSettings.Get("InvocieHeaderDisplayFormLink");
                Invoice._allItemsLink = ConfigurationManager.AppSettings.Get("InvocieAllItemsLink");
                Invoice._divApproverLink = ConfigurationManager.AppSettings.Get("DivisionalApproverViewUrl");
                Invoice._accountingApproverLink = ConfigurationManager.AppSettings.Get("AccountingApproverViewUrl");
                Invoice._InaccountingApproverLink = ConfigurationManager.AppSettings.Get("InAccountingApproverViewUrl");

                Invoice._fromadrress = ConfigurationManager.AppSettings.Get("FromAddress");
                Invoice._toadrress = ConfigurationManager.AppSettings.Get("ToAddress");
                Invoice._subject = ConfigurationManager.AppSettings.Get("Subject");
                Invoice._smtp = ConfigurationManager.AppSettings.Get("SmtpAddress");
                Invoice._smtpUserName = ConfigurationManager.AppSettings.Get("SmtpUserName");
                Invoice._smtpPassword = ConfigurationManager.AppSettings.Get("SmtpPassword");
                if (ConfigurationManager.AppSettings.Get("SmtpPort") != null)
                    Invoice._smtpPort = Convert.ToInt32(ConfigurationManager.AppSettings.Get("SmtpPort"));
                else
                    Invoice._smtpPort = 587;

            }
            catch (Exception ex)
            {
                //log
               
                Logger.log.Error("Error in Reading Configurations "+ex.Message);
                throw ex;
            }
        }
        public static void SendEmail(string body, string To)
        {
            try
            {
                Logger.log.Info("In Sending Email Method");
                MailMessage msg = new MailMessage();
                msg.To.Add(new MailAddress(To));
                msg.From = new MailAddress(Invoice._fromadrress);
                msg.Subject = Invoice._subject;
                msg.Body = body;
                msg.IsBodyHtml = true;
                SmtpClient client = new SmtpClient();
                client.Host = Invoice._smtp;
                // client.UseDefaultCredentials = true;
                client.Credentials = new System.Net.NetworkCredential(Invoice._smtpUserName, Invoice._smtpPassword);
                client.Port = Invoice._smtpPort;
                // client.EnableSsl = true;
                if (Invoice._runforCloud)
                {
                    //client.Port = 465;
                    //client.Port = 587;
                    client.Port = 25;
                    client.EnableSsl = false;
                }
                client.Send(msg);
                Logger.log.Info("Sent Email to : " + To);
            }
            catch (Exception ex)
            {
                Logger.log.Error("Error in Sending Email");
                Console.WriteLine("Error in sending email." + ex.Message);
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
