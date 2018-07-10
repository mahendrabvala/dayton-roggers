using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dayton_RogersTool
{
    class Invoice
    {
        public static string _invoiceSourceLibraryName;
        public static string _invoceDestinationListName;
        public static string _coversheetDetailsListName;
        public static string _siteUrl;
        public static string _domain;
        public static string _userName;
        public static string _password;
        public static string _cheqRegListName;
        public static string _vendorInvoiceListName;
        public static bool _mapVouchers;
        public static string _filePath;
        public static string _repository;
        public static bool _runforCloud;

        //Email Configration
        public static string _fromadrress;
        public static string _toadrress;
        public static string _subject;
        public static string _smtp;
        public static string _smtpUserName;
        public static string _smtpPassword;
        public static int _smtpPort;


        public string InvoiceSourceLibraryName
        {
            get { return _invoiceSourceLibraryName; }
            set { _invoiceSourceLibraryName = value; }
        }
        public string InvoceDestinationListName
        {
            get { return _invoceDestinationListName; }
            set { _invoceDestinationListName = value; }
        }
        public string CoversheetDetailsListName
        {
            get { return _coversheetDetailsListName; }
            set { _coversheetDetailsListName = value; }
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
        public string CheqRegListName
        {
            get { return _cheqRegListName; }
            set { _cheqRegListName = value; }
        }
        public string VendorInvoiceListName
        {
            get { return _vendorInvoiceListName; }
            set { _vendorInvoiceListName = value; }
        }
        public Boolean MapVouchers
        {
            get { return _mapVouchers; }
            set { _mapVouchers = value; }
        }
        public string FilePath
        {
            get { return _filePath; }
            set { _filePath = value; }
        }
        public string Repository
        {
            get { return _repository; }
            set { _repository = value; }
        }
        public Boolean RunForCloud
        {
            get { return _runforCloud; }
            set { _runforCloud = value; }
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
        
    }
}
