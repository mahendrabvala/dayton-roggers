using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GKL.ApprovalEmails
{
    public class Invoice
    {
        public static string _invoiceCoverSheetHeaderListName;
        public static string _orgListName;
        public static string _siteUrl;
        public static string _domain;
        public static string _userName;
        public static string _password;      
        public static bool _runforCloud;
        public static string _displayFormLink;
        public static string _allItemsLink;
        public static string _divApproverLink;
        public static string _accountingApproverLink;
        public static string _InaccountingApproverLink;

        public static string _fromadrress;
        public static string _toadrress;
        public static string _subject;
        public static string _smtp;
        public static string _smtpUserName;
        public static string _smtpPassword;
        public static int _smtpPort;


        public string InvoiceCoverSheetHeadrListName
        {
            get { return _invoiceCoverSheetHeaderListName; }
            set { _invoiceCoverSheetHeaderListName = value; }
        }
        public string OrgListName
        {
            get { return _orgListName; }
            set { _orgListName = value; }
        }

        public string SiteUrl
        {
            get { return _siteUrl; }
            set { _siteUrl = value; }
        }
        public string Domain
        {
            get { return _domain; }
            set { _domain = value; }
        }
        public string UserName
        {
            get { return _userName; }
            set { _userName = value; }
        }
        public string Password
        {
            get { return _password; }
            set { _password = value; }
        }       
        public Boolean RunForCloud
        {
            get { return _runforCloud; }
            set { _runforCloud = value; }
        }

        public string DisplayFormLink
        {
            get { return _displayFormLink; }
            set { _displayFormLink = value; }
        }

        public string AllItemsLink
        {
            get { return _allItemsLink; }
            set { _allItemsLink = value; }
        }

        public string FromAddress
        {
            get { return _fromadrress; }
            set { _fromadrress = value; }
        }

        public string ToAddress
        {
            get { return _toadrress; }
            set { _toadrress = value; }
        }

        public string Subject
        {
            get { return _subject; }
            set { _subject = value; }
        }
        public string SmtpAddress
        {
            get { return _smtp; }
            set { _smtp = value; }
        }

        public string SMTPUserName
        {
            get { return _smtpUserName; }
            set { _smtpUserName = value; }
        }

        public string SMTPPassword
        {
            get { return _smtpPassword; }
            set { _smtpPassword = value; }
        }

        public int SMTPPOrt
        {
            get { return _smtpPort; }
            set { _smtpPort = value; }
        }


        public string DivApproverLink
        {
            get { return _divApproverLink; }
            set { _divApproverLink = value; }
        }

        public string AccountingApproverLink
        {
            get { return _accountingApproverLink; }
            set { _accountingApproverLink = value; }
        }
        public string InAccountingApproverLink
        {
            get { return _InaccountingApproverLink; }
            set { _InaccountingApproverLink = value; }
        }
    }
}
