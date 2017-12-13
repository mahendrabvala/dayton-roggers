using System;
using Microsoft.SharePoint.Client;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Linq;

namespace GKL.ApprovalEmails
{    // testing code in Office 365 GKBLabs
    class Program
    {
        #region variables
        static ClientContext context;
        static DataTable dtHeader;
        static DataTable dtOrgInfo;
        static List listOrganizationInfo = null;
        static StringBuilder body = null;
        private static string companyID;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                //log the start time
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Started Processing..");
                Logger.log.Info("=========== Started Approval Email Notification Service ====================");
                Utils.ReadConfigurationValues();
                dtOrgInfo = ConstructOrgTable();
                dtHeader = ConstructDataTable();
                if (!string.IsNullOrEmpty(Invoice._siteUrl))
                {
                    context = new ClientContext(Invoice._siteUrl);
                }
                else
                {
                    Logger.log.Info("The site url set to null in configuration");
                }
                if (Invoice._runforCloud)
                {
                    Logger.log.Info("The configuration set to run for cloud");
                    System.Security.SecureString pwdSecureString = new System.Security.SecureString();
                    if (!string.IsNullOrEmpty(Invoice._password))
                    {
                        foreach (char c in Invoice._password)
                        {
                            pwdSecureString.AppendChar(c);
                        }
                        context.Credentials = new SharePointOnlineCredentials(Invoice._userName, pwdSecureString);
                    }
                }
                else
                {
                    Logger.log.Info("The configuration set to run for On Premise environment");
                    context.Credentials = new NetworkCredential(Invoice._userName, Invoice._password, Invoice._domain);
                }

                Web web = context.Web;
                context.Load(web, w => w.Title, w => w.Url);
                List listHeader = web.Lists.GetByTitle(Invoice._invoiceCoverSheetHeaderListName);
                listOrganizationInfo = web.Lists.GetByTitle(Invoice._orgListName);

                //Updating the Status from Dvisional Approval Approved to Accounting Approval Pending
                UpdateDivApprovedtoAccountPendingStatus(listHeader);
                Console.WriteLine("Status Update Done.");

                dtHeader = GetHeaderInfo(listHeader);

                DataTable dtCompanyIDs = dtHeader.DefaultView.ToTable(true, "InvoiceCompany");
                if (dtCompanyIDs != null)
                {
                    foreach (DataRow dr in dtCompanyIDs.Rows)
                    {
                        string companyID = dr["InvoiceCompany"].ToString();

                        // Divisional Approval Pending
                        try
                        {
                            DataRow[] drDvisions = dtHeader.Select("InvoiceCompany=" + companyID + "AND DivisionalApproval = 'Divisional Approval Pending'");
                            string divApprovalType = "divisional";
                            if (drDvisions.Length != 0)
                            {
                                Logger.log.Info("Sending Mails to Divisional Approvers ");
                                StringBuilder divEmailBuilder = ConstructEmailBody(drDvisions, divApprovalType);
                                SendApproverEmailRejected(companyID, divApprovalType, divEmailBuilder);
                                Console.WriteLine("Sent Emails to Divisional Approvers for Company " + companyID);
                            }
                        }
                        catch (Exception ex)
                        {
                           // Logger.log.Error("Error in Getting Divisional Approvers Details " + ex.ToString());
                          //  Console.WriteLine("Error {0}", ex);
                        }

                        //Accounting Approval Pending
                        try
                        {
                            DataRow[] drAccounts = dtHeader.Select("InvoiceCompany=" + companyID + "AND AccountingApproval ='Accounting Approval Pending'");
                            string accntApprovalType = "accounting";
                            if (drAccounts.Length != 0)
                            {
                                Logger.log.Info("Sending Mails to Accounting Approvers Rejected");
                                StringBuilder accntEmailBuilder = ConstructEmailBody(drAccounts, accntApprovalType);
                                SendApproverEmailRejected(companyID, accntApprovalType, accntEmailBuilder);
                                Console.WriteLine("Sent Emails to Accounting Approvers for Company " + companyID);
                            }
                        }
                        catch (Exception ex)
                        {
                           // Logger.log.Error("Error in Getting Accounting Approvers Details " + ex.ToString());
                        }
                    }

                    //  Divisional Approval Rejected
                    try
                    {
                        DataRow[] drDvisions = dtHeader.Select("DivisionalApproval = 'Divisional Approval Rejected'");
                        string divApprovalType = "divisional";
                        if (drDvisions.Length != 0)
                        {
                            Logger.log.Info("Sending Mails to Divisional Approvers Rejected ");
                            StringBuilder divEmailBuilder = ConstructEmailBodyRejected(drDvisions, divApprovalType);
                            SendApproverEmailRejected(divApprovalType, divEmailBuilder);
                            Console.WriteLine("Sent Emails to  Divisional Approval Rejected");
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.log.Error("Error in Getting Divisional Approvers Details " + ex.ToString());
                        Console.WriteLine("Error {0}", ex);
                    }

                    // Accounting Approval Rejected
                    try
                    {
                        DataRow[] drAccounts = dtHeader.Select("AccountingApproval ='Accounting Approval Rejected'");
                        string accntApprovalType = "accounting";
                        if (drAccounts.Length != 0)
                        {
                            Logger.log.Info("Sending Mails to Accounting Approvers");
                            StringBuilder accntEmailBuilder = ConstructEmailBodyRejected(drAccounts, accntApprovalType);
                            SendApproverEmailRejected(accntApprovalType, accntEmailBuilder);
                            Console.WriteLine("Sent Emails to Accounting Approvers Rejected");
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.log.Error("Error in Getting Accounting Approvers Details " + ex.ToString());
                    }
                }
                Logger.log.Info("=========== Completed Approval Email Notification Service ====================");
                Console.WriteLine("Process completed, please close the window to exit process..!");
                //Console.ReadKey();
            }
            catch (Exception ex)
            {
                Logger.log.Error("Error in Approval Email " + ex.ToString());
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error in sending Invoice approval mails, Please check all the configurations and try again..{0}", ex);
                Console.ResetColor();
               // Console.ReadLine();
            }
        }

        private static void SendApproverEmailRejected(string companyID, string approvalType, StringBuilder emailBody)
        {
           // Utils.SendEmail("Testing", "accounting@daytonrogers.com");
            Logger.log.Info("Getting Company "+companyID+" Approvers");
            //orglist
            CamlQuery camlOrgQuery = new CamlQuery();
            camlOrgQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Company_x0020_ID'/>" +
            "<Value Type='Text'>" + companyID + "</Value></Eq></Where></Query><RowLimit>1000</RowLimit></View>";
            ListItemCollection orgInfoCollection = listOrganizationInfo.GetItems(camlOrgQuery);
            context.Load(orgInfoCollection, items => items.Include(
                     item => item.Id,
                     item => item["Title"],
                      item => item["Company_x0020_ID"],
                       item => item["Divisional_x0020_Approvers"],
                        item => item["Accounting_x0020_Approvers"]
                     ));
            context.ExecuteQuery();
            if (orgInfoCollection.Count > 0)
            {
              //  Console.WriteLine("Email Id", orgInfoCollection.Context.Credentials);
                foreach (ListItem orgItem in orgInfoCollection)
                {
                    if (approvalType.ToLower() == "divisional")
                    {
                        if (orgItem["Divisional_x0020_Approvers"] != null)
                        {
                            FieldUserValue[] fuvDA = (FieldUserValue[])orgItem["Divisional_x0020_Approvers"];
                            User divisionalApprover = null;
                            for (int i = 0; i < fuvDA.Length; i++)
                            {
                                divisionalApprover = context.Web.EnsureUser(fuvDA[i].LookupValue);
                                context.Load(divisionalApprover);
                                context.ExecuteQuery();
                                string divApproverLoginName = divisionalApprover.LoginName;
                                string divApproverEmailID = divisionalApprover.Email;
                                if (!string.IsNullOrEmpty(divApproverEmailID))
                                {
                                    string body = emailBody.ToString().Replace("[UserName]", divisionalApprover.Title);
                                    Utils.SendEmail(body, divApproverEmailID);

                                }

                                else
                                {

                                    divApproverLoginName = divApproverLoginName.Split('|').Last();
                                    string body = emailBody.ToString().Replace("[UserName]", divisionalApprover.Title);
                                    Utils.SendEmail(body, divApproverLoginName);

                                }

                            }
                        }
                    }
                    else
                    {
                        if (orgItem["Accounting_x0020_Approvers"] != null)
                        {
                            FieldUserValue[] fuvAccnt = (FieldUserValue[])orgItem["Accounting_x0020_Approvers"];
                            User accntApprover = null;
                            for (int i = 0; i < fuvAccnt.Length; i++)
                            {
                                accntApprover = context.Web.EnsureUser(fuvAccnt[i].LookupValue);
                                context.Load(accntApprover);
                                context.ExecuteQuery();
                                string accountingApproverLoginName = accntApprover.LoginName;
                                string accountingApproverEmailID = accntApprover.Email;
                                if (!string.IsNullOrEmpty(accountingApproverEmailID))
                                {
                                    string body = emailBody.ToString().Replace("[UserName]", accntApprover.Title);
                                    Utils.SendEmail(body, accountingApproverEmailID);
                                }
                                else
                                {

                                    accountingApproverLoginName = accountingApproverLoginName.Split('|').Last();
                                    string body = emailBody.ToString().Replace("[UserName]", accntApprover.Title);
                                    Utils.SendEmail(body, accountingApproverLoginName);

                                }
                            }
                        }
                    }
                }
            }
            
        }

        private static void SendApproverEmailRejected(string approvalType, StringBuilder emailBody)
        {
            //Accounting
            try
            {
                string accountingApproverEmailID = "accounting@daytonrogers.com";
                
                    string body = emailBody.ToString();
                    Utils.SendEmail(body, accountingApproverEmailID);
                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Email Error :", ex);
                Console.ReadKey();
            }
        }

        private static DataTable GetHeaderInfo(List listHeader)
        {
            Logger.log.Info("Constructing Invoice Header Table");
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                "<Value Type='Number'>1</Value></Geq></Where></Query></View>";
            ListItemCollection headerItemCollection = listHeader.GetItems(camlQuery);
            context.Load(headerItemCollection, items => items.Include(
                     item => item.Id,
                     item => item["Title"],
                      item => item["Vendor_x0020_Name"],
                       item => item["Vendor_x0020_Number"],
                        item => item["PO_x0020_Number"],
                         item => item["Invoice_x0020_Company"],
                          item => item["Invoice_x0020_Date"],
                           item => item["Invoice_x0020_Amount"],
                           item => item["Invoice_x0020_Division"],
                     item => item["Divisional_x0020_Approval"],
                        item => item["Accounting_x0020_Approval"],
                     item => item["FileRef"]
                     ));
            context.ExecuteQuery();
            dtHeader.Clear();
            foreach (ListItem item in headerItemCollection)
            {
                DataRow dr = dtHeader.NewRow();
                dr["ID"] = item.Id.ToString();
                dr["Title"] = item["Title"].ToString();
                if (item["Vendor_x0020_Name"] != null)
                    dr["VendorName"] = item["Vendor_x0020_Name"].ToString();
                if (item["Vendor_x0020_Number"] != null)
                    dr["Vendor"] = item["Vendor_x0020_Number"].ToString();
                if (item["PO_x0020_Number"] != null)
                    dr["PO"] = item["PO_x0020_Number"].ToString();
                if (item["Invoice_x0020_Company"] != null)
                    dr["InvoiceCompany"] = item["Invoice_x0020_Company"].ToString();
                if (item["Invoice_x0020_Date"] != null)
                    dr["InvoiceDate"] = item["Invoice_x0020_Date"].ToString();
                if (item["Invoice_x0020_Amount"] != null)
                    dr["InvoiceAmount"] = item["Invoice_x0020_Amount"].ToString();
                if (item["Invoice_x0020_Division"] != null)
                    dr["InvoiceDivision"] = item["Invoice_x0020_Division"].ToString();
                if (item["Divisional_x0020_Approval"] != null)
                    dr["DivisionalApproval"] = item["Divisional_x0020_Approval"].ToString();
                if (item["Accounting_x0020_Approval"] != null)
                    dr["AccountingApproval"] = item["Accounting_x0020_Approval"].ToString();

                dr["LinktoItem"] = Invoice._displayFormLink + dr["ID"].ToString(); //String.Format("{0}/{1}?ID={2}", web.Url, item.ParentList.Forms[4].ServerRelativeUrl, item.Id);
                dtHeader.Rows.Add(dr);
            }
            return dtHeader;
        }
        private static void UpdateDivApprovedtoAccountPendingStatus(List listHeader)
        {
            Logger.log.Info("Process started to Update status from Divisional Approval Approved to Accounting Approval Pending");

            CamlQuery camlStatusUpdateQuery = new CamlQuery();
            camlStatusUpdateQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Divisional_x0020_Approval'/>" +
                "<Value Type='Choice'>Divisional Approval Approved</Value></Eq></Where></Query></View>";
            ListItemCollection divApprovedCollection = listHeader.GetItems(camlStatusUpdateQuery);
            context.Load(divApprovedCollection, items => items.Include(
                     item => item.Id,
                     item => item["Title"],
                     item => item["Divisional_x0020_Approval"],
                        item => item["Accounting_x0020_Approval"],
                     item => item["FileRef"]
                     ));
            context.ExecuteQuery();
            Logger.log.Info("Got "+ divApprovedCollection.Count+" Items with  Divisional Approval Approved Status ");

            foreach (ListItem item in divApprovedCollection)
            {
                if (item["Divisional_x0020_Approval"] != null && item["Divisional_x0020_Approval"].ToString() == "Divisional Approval Approved")
                {
                    if (item["Accounting_x0020_Approval"] != null && item["Accounting_x0020_Approval"].ToString() == "None")
                    {
                        Console.WriteLine("Updating Item " + item.Id);
                        Logger.log.Info("Updating the status for the item : " + item.Id); 
                        item["Accounting_x0020_Approval"] = "Accounting Approval Pending";
                        item.Update();
                        context.Load(item);
                    }
                    else
                    {
                        Logger.log.Info("The Accounting Approval Status is Not None for the Item : "+item.Id +" hence not changing the status "); ;
                    }
                }
            }
            context.ExecuteQuery();
            Logger.log.Info("Status changed from Divisional Approval Approved to Accounting Approval Pending for all the eligible items in header"); 
        }
        private static StringBuilder ConstructEmailBody(DataRow[] drCollection, string approvalType)
        {

            Logger.log.Info("Constructing Email Body with "+drCollection.Length+" items for  "+approvalType);

            string companyID = string.Empty; //Used for view link of mail.
            body = new StringBuilder();
            body.Append("Hi [UserName] ,");
            //   Console.WriteLine("User{0}",web.CurrentUser));
            // Console.ReadKey());
            body.Append("<br>");
            body.Append("<P>");
            body.Append("The following invoices are uploaded and awaiting your Approval [ " + approvalType + " ] : " + "<br/>");
            body.Append("</p>");
            body.Append("<table border=1 style='width: 100 %; border - style: solid; border - width: 1px; '>");
            body.Append("<tr style='background-color:#7FFFD4' > ");
            body.Append("<td  > Link </td>");
            body.Append("<td  > ID</td>");
            body.Append("<td  > Invoice Number</td>");
            body.Append("<td  > Vendor Name</td>");
            body.Append("<td  > Vendor No</td>");
            body.Append("<td  > PO No</td>");
            body.Append("<td  > Invoice Company</td>");
            body.Append("<td  > Invoice Date</td>");
            body.Append("<td  > Invoice Amount</td>");
            body.Append("<td  > Invoice Division</td>");
            body.Append("<td  > Payment Status</td>");
            body.Append("</tr>");

            foreach (DataRow drOrgItem in drCollection)
            {
               
                string displayFormLink = "<a href='" + Invoice._displayFormLink + drOrgItem["ID"].ToString() + "'>View</a>";
                body.Append("<tr>");
                body.Append("<td >" + displayFormLink + "</td>");
                body.Append("<td >" + drOrgItem["ID"].ToString() + "</td>");
                body.Append("<td >" + drOrgItem["Title"].ToString() + "</td>");
                if (drOrgItem["VendorName"] != null)
                    body.Append("<td  >" + drOrgItem["VendorName"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["Vendor"] != null)
                    body.Append("<td >" + drOrgItem["Vendor"].ToString() + " </td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");

                if (drOrgItem["PO"] != null)
                    body.Append("<td >" + drOrgItem["PO"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceCompany"] != null)
                {
                    companyID = drOrgItem["InvoiceCompany"].ToString();
                    body.Append("<td >" + drOrgItem["InvoiceCompany"].ToString() + "</td>");
                }
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceDate"] != null)
                    body.Append("<td >" + drOrgItem["InvoiceDate"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceAmount"] != null)
                    body.Append("<td >" + drOrgItem["InvoiceAmount"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceDivision"] != null)
                    body.Append("<td >" + drOrgItem["InvoiceDivision"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (approvalType == "divisional")
                {
                    if (drOrgItem["DivisionalApproval"] != null)
                        body.Append("<td >" + drOrgItem["DivisionalApproval"].ToString() + "</td>");
                    else
                        body.Append("<td  >" + string.Empty + "</td>");
                }
                else
                {
                    if (drOrgItem["AccountingApproval"] != null)
                        body.Append("<td >" + drOrgItem["AccountingApproval"].ToString() + "</td>");
                    else
                        body.Append("<td  >" + string.Empty + "</td>");
                }

                body.Append("</tr>");
            }
            body.Append("</table>");
            body.Append("<br>");
            body.Append("<table border=1 style='width: 100 %; border - style: solid); border - width: 5px; '>");
            body.Append("<tr>");
            if (approvalType == "divisional")
            {
                body.Append("<a href='" + Invoice._divApproverLink + companyID+ "'>Click here to view all Invoices</a>" + "</ br>");
            }
            else if (approvalType == "accounting")
            {
                body.Append("<a href='" + Invoice._accountingApproverLink + companyID+ "'>Click here to view all Invoices</a>" + "</ br>");
                body.Append("<br>");
                body.Append("<a href='" + Invoice._InaccountingApproverLink + companyID + "'>Click here to view all In Progress Accounting Invoices in Cover Sheet Details</a>" + "</ br>");
            }
            else
                body.Append("<a href='" + Invoice._allItemsLink + "'>Click here to view all Invoices</a>" + "</ br>");
            body.Append("</tr>");
            body.Append("<P>");
            body.Append("Thanks");
            body.Append("<br>");
            body.Append("Admin.</ br>");
            body.Append("</P>");
            body.Append("<P>");
            body.Append("<b> Note:</b><i>" + " This is an automatically generated email from Dayton Rogers SharePoint Server." + "</i>");
            body.Append("</P>");

            return body;
        }


        //  Divisional Accounting Method Approval Rejected


        private static StringBuilder ConstructEmailBodyRejected(DataRow[] drCollection, string approvalType)
        {

            Logger.log.Info("Constructing Email Body with " + drCollection.Length + " items for  " + approvalType);

            string companyID = string.Empty; //Used for view link of mail.
            body = new StringBuilder();
            body.Append("Hi Accounting ,");
            //   Console.WriteLine("User{0}",web.CurrentUser));
            // Console.ReadKey());
            body.Append("<br>");
            body.Append("<P>");
            body.Append("The following invoices are uploaded and awaiting your Approval [ " + approvalType + " ] : " + "<br/>");
            body.Append("</p>");
            body.Append("<table border=1 style='width: 100 %; border - style: solid; border - width: 1px; '>");
            body.Append("<tr style='background-color:#7FFFD4' > ");
           // body.Append("<td  > Link </td>");
          //  body.Append("<td  > ID</td>");
            body.Append("<td  > Invoice Number</td>");
            body.Append("<td  > Vendor Name</td>");
            body.Append("<td  > Vendor No</td>");
            body.Append("<td  > PO No</td>");
            body.Append("<td  > Invoice Company</td>");
            body.Append("<td  > Invoice Date</td>");
            body.Append("<td  > Invoice Amount</td>");
            body.Append("<td  > Invoice Division</td>");
            body.Append("<td  > Payment Status</td>");
            body.Append("</tr>");

            foreach (DataRow drOrgItem in drCollection)
            {

                string displayFormLink = "<a href='" + Invoice._displayFormLink + drOrgItem["ID"].ToString() + "'>View</a>";
                body.Append("<tr>");
            //    body.Append("<td >" + displayFormLink + "</td>");
              //  body.Append("<td >" + drOrgItem["ID"].ToString() + "</td>");
                body.Append("<td >" + drOrgItem["Title"].ToString() + "</td>");
                if (drOrgItem["VendorName"] != null)
                    body.Append("<td  >" + drOrgItem["VendorName"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["Vendor"] != null)
                    body.Append("<td >" + drOrgItem["Vendor"].ToString() + " </td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");

                if (drOrgItem["PO"] != null)
                    body.Append("<td >" + drOrgItem["PO"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceCompany"] != null)
                {
                    companyID = drOrgItem["InvoiceCompany"].ToString();
                    body.Append("<td >" + drOrgItem["InvoiceCompany"].ToString() + "</td>");
                }
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceDate"] != null)
                    body.Append("<td >" + drOrgItem["InvoiceDate"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceAmount"] != null)
                    body.Append("<td >" + drOrgItem["InvoiceAmount"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (drOrgItem["InvoiceDivision"] != null)
                    body.Append("<td >" + drOrgItem["InvoiceDivision"].ToString() + "</td>");
                else
                    body.Append("<td  >" + string.Empty + "</td>");
                if (approvalType == "divisional")
                {
                    if (drOrgItem["DivisionalApproval"] != null)
                        body.Append("<td >" + drOrgItem["DivisionalApproval"].ToString() + "</td>");
                    else
                        body.Append("<td  >" + string.Empty + "</td>");
                }
                else
                {
                    if (drOrgItem["AccountingApproval"] != null)
                        body.Append("<td >" + drOrgItem["AccountingApproval"].ToString() + "</td>");
                    else
                        body.Append("<td  >" + string.Empty + "</td>");
                }

                body.Append("</tr>");
            }
            body.Append("</table>");
            body.Append("<br>");
            body.Append("<P>");
            body.Append("Thanks");
            body.Append("<br>");
            body.Append("Admin.</ br>");
            body.Append("</P>");
            body.Append("<P>");
            body.Append("<b> Note:</b><i>" + " This is an automatically generated email from Dayton Rogers SharePoint Server." + "</i>");
            body.Append("</P>");

            return body;
        }


        private static DataTable ConstructOrgTable()
        {

            Logger.log.Info("Constructing Organization Table");
            dtOrgInfo = new DataTable();
            dtOrgInfo.Columns.Add("ID");
            dtOrgInfo.Columns.Add("Title");
            dtOrgInfo.Columns.Add("CompanyID");
            dtOrgInfo.Columns.Add("DivisionApprovers");
            dtOrgInfo.Columns.Add("AccountingApprovers");
            return dtOrgInfo;
        }

        private static DataTable ConstructDataTable()
        {
            Logger.log.Info("Constructing Invoice Header Table");
            dtHeader = new DataTable();
            dtHeader.Columns.Add("ID");
            dtHeader.Columns.Add("Title");
            dtHeader.Columns.Add("VendorName");
            dtHeader.Columns.Add("Vendor");
            dtHeader.Columns.Add("PO");
            dtHeader.Columns.Add("InvoiceDate");
            dtHeader.Columns.Add("InvoiceCompany");
            dtHeader.Columns.Add("InvoiceDivision");
            dtHeader.Columns.Add("InvoiceAmount");
            dtHeader.Columns.Add("DivisionalApproval");
            dtHeader.Columns.Add("AccountingApproval");
            dtHeader.Columns.Add("LinktoItem");
            dtHeader.Columns.Add("Company");
            //add another column here
            return dtHeader;
        }
    }
}