using System;
using Microsoft.SharePoint.Client;
using System.Data;
using System.Net;
using System.IO;
using log4net;
using System.Diagnostics;
using System.Collections;

namespace Dayton_RogersTool
{
    class Program
    {
        #region variables
        static ClientContext context;
        static DataTable dtInvoices;
        public static string BrokenInvoicesLog = string.Empty;
        public static string ErrorMessage = string.Empty;
        static List listCoversheet = null;
        static int processedCount = 0;
        #endregion

        static void Main(string[] args)
        {
            try
            {
                #region InvoiceSplit
                //log the start time
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Started Processing..");
                Logger.log.Info("=========== Started Invoice Processing ====================");
                BrokenInvoicesLog = BrokenInvoicesLog + "<p>Hi Admin, </p>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<p>Coversheet Processing Status Report: <strong>" + DateTime.Now.ToString("MMMM") + " " + DateTime.Now.Day + ", " + DateTime.Now.Year + "</strong></p>";

                BrokenInvoicesLog = BrokenInvoicesLog + "<table border=1 style='width: 100 %; border - style: solid; border - width: 1px;' > ";
                BrokenInvoicesLog = BrokenInvoicesLog + "<tr style='background-color:#7FFFD4' >";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > ID </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Invoice # </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Vendor Name </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Vendor No </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Invoice Company </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Invoice Date </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Invoice Amount </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Status </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "<th  > Error </th>";
                BrokenInvoicesLog = BrokenInvoicesLog + "</tr>";

                Utils.ReadConfigurationValues();

                dtInvoices = ConstructDataTable();
                if (!string.IsNullOrEmpty(Invoice._siteUrl)) {
                    context = new ClientContext(Invoice._siteUrl);
                } else {
                    Logger.log.Info("The site url is empty in configuration.");
                }
                if (Invoice._runforCloud) {
                    Logger.log.Info("Configuration set to run for Sharepoint Online.");
                    System.Security.SecureString pwdSecureString = new System.Security.SecureString();
                    if (!string.IsNullOrEmpty(Invoice._password))
                    {
                        foreach (char c in Invoice._password)
                        {
                            pwdSecureString.AppendChar(c);
                        }
                        context.Credentials = new SharePointOnlineCredentials(Invoice._userName, pwdSecureString);
                    }
                } else {
                    Logger.log.Info("Configuration set to run for SharePoint On Premise.");
                    context.Credentials = new NetworkCredential(Invoice._userName, Invoice._password, Invoice._domain);
                }

                Web web = context.Web;
                context.ExecuteQuery();
                Logger.log.Info("Context created for given site " + Invoice._siteUrl);

                listCoversheet = web.Lists.GetByTitle(Invoice._invoiceSourceLibraryName);
                List listHeader = web.Lists.GetByTitle(Invoice._invoceDestinationListName);
                List listDetails = web.Lists.GetByTitle(Invoice._coversheetDetailsListName);//To push

                Logger.log.Info("Reading items from List : " + Invoice._invoiceSourceLibraryName);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='processed'/><Value Type='Choice'>False</Value></Eq></Where></Query></View>";

                ListItemCollection coversheetCollection = listCoversheet.GetItems(camlQuery);

                Logger.log.Info("CamlQuery: " + camlQuery.ViewXml);

                context.Load(
                    coversheetCollection,
                    items => items.Include(
                    item => item.Id,
                    item => item["Vendor_x0020_Name"],
                    item => item["Vendor"],
                    item => item["PO"],
                    item => item["Invoice_x0020_Co"],
                    item => item["Invoice_x0020_DIV"],
                    item => item["Invoice_x0020__x0023_"],//Invoice#
                    item => item["Invoice_x0020_Date"],
                    item => item["Invoice_x0020_Amt"],
                    item => item["Company"],
                    item => item["Division"],
                    item => item["Dept_x002d_3_x0020_CH"],
                    item => item["Account_x002d_5_x0020_CH"],
                    item => item["Expense_x0020_Amt"],
                    item => item["Job"],
                    item => item["CER"],
                    //item => item["Monetary_Unit"],
                    item => item["Acct_x0020_Date"],
                    item => item["processed"],
                    item => item["Account_x0020_String"]
                ));
                context.ExecuteQuery();
                dtInvoices.Clear();

                //"Filling datatable with CS Invoice data if invoice alreadey not exist in header list;
                FillCoversheetInfotoDataTable(dtInvoices, coversheetCollection);

                //Console.WriteLine("Invoice Processing Completed!");
                if (dtInvoices != null && dtInvoices.Rows.Count > 0) {
                    Logger.log.Info("Retrieved " + dtInvoices.Rows.Count + " Invoices from CoverSheet ");
                    PushtoHeaderList(dtInvoices, listHeader, listDetails);
                } else {
                    Logger.log.Info("Retrieved " + dtInvoices.Rows.Count + " Invoices from CoverSheet ");
                }
                Console.WriteLine("Invoice Processing Completed!");
                Console.WriteLine("Splitting is Completed.");
                Logger.log.Info("Splitting Process is Completed.. ");
                if (Invoice._mapVouchers) {
                    Logger.log.Info("Starting Mapping the Voucher Numbers");
                    Console.WriteLine("Started Mapping Voucher Number!");

                    if (!string.IsNullOrEmpty(Invoice._filePath)) {
                        string filepath = string.Empty;
                        if (System.IO.File.Exists(Invoice._filePath)) {
                            Logger.log.Info("Getting voucher mappings file from filepath");
                            filepath = Invoice._filePath;
                        } else {
                            Logger.log.Info("Getting voucher mappings file from folder path");
                            Console.WriteLine("Getting file from folder  :" + Invoice._filePath);
                            filepath = GetFileFromFolderPath(Invoice._filePath);
                        }
                        if (!string.IsNullOrEmpty(filepath)) {

                            string destinationPath = string.Empty;
                            if (System.IO.File.Exists(Invoice._repository)) {
                                Logger.log.Info("Reading destination mapping vouchers filepath ");
                                destinationPath = Invoice._repository;
                            } else {
                                Logger.log.Info("Reading destination mapping vouchers filepath from folder");
                                string fileName = Path.GetFileName(filepath);
                                if (Invoice._repository.ToString().ToLower().Contains(fileName.ToLower()))
                                {
                                    destinationPath = Invoice._repository;
                                }
                                else
                                {
                                    destinationPath = Invoice._repository + fileName;
                                }
                            }

                            Logger.log.Info("Reading voucher map data from given file " + filepath);
                            DataTable dtVouchers = new DataTable();
                            /*dtVouchers = Utils.ReadExcelData(filepath);
                            bool mappingCompleted = MapVoucher(dtVouchers, listHeader);
                            try
                            {
                                if (mappingCompleted)
                                {
                                    Logger.log.Info("Mapping is completed.");
                                    MoveFile(filepath, destinationPath);
                                    Console.WriteLine("Voucher Mapping Process Completed.!");
                                }
                            }
                            catch (Exception)
                            {
                                Console.WriteLine("Error in moving the file");
                            }*/
                        } else {
                            Console.WriteLine("Unable to map the voucher, reason could be the file path is not correct or the file doenst exist in the specified locaton : " + filepath);
                            Logger.log.Info("Unable to map the voucher, reason could be the file path is not correct or the file doenst exist in the specified locaton : " + filepath);
                        }
                    }
                } else {
                    Logger.log.Info("Mapping Voucher set to false in configuration.");
                }
                Console.WriteLine("Please Close the Window to exit Process..!");
                #endregion InvoiceSplit
                if (!string.IsNullOrEmpty(BrokenInvoicesLog)) {
                    BrokenInvoicesLog = BrokenInvoicesLog + " </table><br />";
                    string LogFileName = DateTime.Now.ToString("yyyyMMdd");
                    BrokenInvoicesLog = BrokenInvoicesLog + "<p>Complete log file: C:\\Dayton Rogers Log History\\InvoiceSplit_" + LogFileName + ".log </p>";
                    BrokenInvoicesLog = BrokenInvoicesLog + "<p>Thank you</p>";
                    Utils.SendEmail(BrokenInvoicesLog, Invoice._toadrress);
                }
                Logger.log.Info("============= Completed Invoice Processing ====================");
                //Console.ReadLine();
            } catch (Exception ex) {
                //Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Error in Processing Invoice, Please check all the configurations and try again..{0}", ex);
                Logger.log.Error("Error in Processing Invoices, Message : " + ex.ToString());
                try
                {
                    if (!string.IsNullOrEmpty(BrokenInvoicesLog)) {
                        //BrokenInvoices Email Body is not Empty then send Email
                        BrokenInvoicesLog = BrokenInvoicesLog + " </table><br />";
                        string LogFileName = DateTime.Now.ToString("yyyyMMdd");
                        BrokenInvoicesLog = BrokenInvoicesLog + "<p>Complete log file: C:\\Dayton Rogers Log History\\InvoiceSplit_" + LogFileName + ".log </p>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<p>" + " <b> Error in Processing Invoice </b>  " + ex.ToString() + "</p>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<p>Thank you</p>";
                        Utils.SendEmail(BrokenInvoicesLog, Invoice._toadrress);
                    }
                }
                catch (Exception)
                {
                    //Leave
                }
            }
            Console.ReadLine();
        }

        private static DataTable FillCoversheetInfotoDataTable(DataTable dtInvoices, ListItemCollection coversheetCollection)
        {
            Logger.log.Info("Filling Coversheet records to DataTable");
            foreach (ListItem item in coversheetCollection) {
                if (item["Invoice_x0020__x0023_"] != null) {

                    DataRow dr = dtInvoices.NewRow();
                    dr["ID"] = Convert.ToString(item.Id);
                    if (item["Vendor_x0020_Name"] != null)
                        dr["VendorName"] = item["Vendor_x0020_Name"].ToString();
                    else
                        dr["VendorName"] = string.Empty;

                    if (item["Vendor"] != null)
                        dr["Vendor"] = item["Vendor"].ToString();
                    else
                        dr["Vendor"] = string.Empty;

                    if (item["PO"] != null)
                        dr["PO"] = item["PO"].ToString();
                    else
                        dr["PO"] = string.Empty;

                    if (item["Invoice_x0020__x0023_"] != null)
                        dr["InvoiceNumber"] = item["Invoice_x0020__x0023_"].ToString();
                    else
                        dr["InvoiceNumber"] = string.Empty;

                    if (item["Invoice_x0020_Date"] != null)
                        dr["InvoiceDate"] = item["Invoice_x0020_Date"].ToString();
                    else
                        dr["InvoiceDate"] = string.Empty;

                    if (item["Invoice_x0020_Co"] != null)
                        dr["InvoiceCompany"] = item["Invoice_x0020_Co"].ToString();
                    else
                        dr["InvoiceCompany"] = string.Empty;

                    if (item["Invoice_x0020_DIV"] != null)
                        dr["InvoiceDivision"] = item["Invoice_x0020_DIV"].ToString();
                    else
                        dr["InvoiceDivision"] = string.Empty;

                    if (item["Invoice_x0020_Amt"] != null)
                        dr["InvoiceAmount"] = item["Invoice_x0020_Amt"].ToString();
                    else
                        dr["InvoiceAmount"] = string.Empty;

                    if (item["Company"] != null)
                        dr["Company"] = item["Company"].ToString();
                    else
                        dr["Company"] = string.Empty;

                    if (item["Division"] != null)
                        dr["Division"] = item["Division"].ToString();
                    else
                        dr["Division"] = string.Empty;

                    if (item["Dept_x002d_3_x0020_CH"] != null)
                        dr["Dept3CH"] = item["Dept_x002d_3_x0020_CH"].ToString();
                    else
                        dr["Dept3CH"] = string.Empty;

                    if (item["Account_x002d_5_x0020_CH"] != null)
                        dr["Account5CH"] = item["Account_x002d_5_x0020_CH"].ToString();
                    else
                        dr["Account5CH"] = string.Empty;

                    if (item["Expense_x0020_Amt"] != null)
                        dr["ExpenseAmt"] = item["Expense_x0020_Amt"].ToString();
                    else
                        dr["ExpenseAmt"] = string.Empty;

                    if (item["Job"] != null)
                        dr["Job"] = item["Job"].ToString();
                    else
                        dr["Job"] = string.Empty;

                    if (item["CER"] != null)
                        dr["CER"] = item["CER"].ToString();
                    else
                        dr["CER"] = string.Empty;

                    // if (item["Monetary_Unit"] != null)
                    //     dr["MonetaryUnit"] = item["Monetary_Unit"].ToString();
                    if (item["Acct_x0020_Date"] != null)
                        dr["AccntDate"] = item["Acct_x0020_Date"].ToString();
                    else
                        dr["AccntDate"] = string.Empty;

                    if (item["Account_x0020_String"] != null)
                        dr["AccountString"] = item["Account_x0020_String"].ToString();
                    else
                        dr["AccountString"] = string.Empty;

                    dtInvoices.Rows.Add(dr);
                }
            }

            return dtInvoices;
        }
        private static void PushtoHeaderList(DataTable dtInvoices, List listHeader, List listDetails)
        {
            Logger.log.Info("Pushing Invoice data to header list");
            if (listHeader != null) {
                foreach (DataRow dr in dtInvoices.Rows) {
                    BrokenInvoicesLog = BrokenInvoicesLog + "<tr>";
                    ErrorMessage = "";
                    try
                    {
                        if (dr["InvoiceNumber"] != null) { //malli added dr["Account5CH"] != DBNull.Value

                            Logger.log.Info("Processing Invoice " + dr["InvoiceNumber"].ToString());
                            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                            ListItem newItem = listHeader.AddItem(itemCreateInfo);
                            newItem["Title"] = Convert.ToString(dr["InvoiceNumber"]);//vendor invoice number
                            newItem["Vendor_x0020_Name"] = Convert.ToString(dr["VendorName"]);
                            newItem["Vendor_x0020_Number"] = Convert.ToString(dr["Vendor"]);
                            newItem["PO_x0020_Number"] = Convert.ToString(dr["PO"]);
                            newItem["Invoice_x0020_Company"] = Convert.ToString(dr["InvoiceCompany"]);
                            newItem["Invoice_x0020_Division"] = Convert.ToString(dr["InvoiceDivision"]);
                            newItem["Invoice_x0020_Date"] = Convert.ToString(dr["InvoiceDate"]);
                            newItem["Invoice_x0020_Amount"] = Convert.ToString(dr["InvoiceAmount"]);
                            newItem.Update();
                            context.ExecuteQuery();
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > " + Convert.ToString(newItem.Id) + " </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > " + Convert.ToString(dr["InvoiceNumber"]) + " </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["VendorName"])+" </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["Vendor"])+" </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["InvoiceCompany"])+" </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["InvoiceDate"])+" </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["InvoiceAmount"])+" </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > Success </td>";
                           
                            Logger.log.Info("Invoice " + Convert.ToString(dr["InvoiceNumber"]) + " Pushed to Header List");
                            string sourceItemId = Convert.ToString(newItem.Id);
                            //Pushing to Details List and storing tde push status in to a string
                            string detailsSaveResult = PushtoDetailsList(dr, listDetails, sourceItemId);
                            BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+ ErrorMessage +" </td>";
                            BrokenInvoicesLog = BrokenInvoicesLog + "</tr>";
                            //Updating Coversheet "Processed" column witd tde push status
                            UpdateCSInvoiceStatus(Convert.ToInt32(dr["ID"].ToString()), dr["InvoiceNumber"].ToString(), detailsSaveResult);
                            Logger.log.Info("Updating Invoice " + dr["InvoiceNumber"].ToString() + " as " + detailsSaveResult + " Invoices");

                        }
                    } catch (Exception ex) {
                        Logger.log.Info("Error in Pushing the Header List Record for the Invoice " + dr["InvoiceNumber"].ToString() + "Error " + ex.Message);
                        //If any exception in pushing to header, make the coversheet processed column to Error
                        UpdateCSInvoiceStatus(Convert.ToInt32(dr["ID"].ToString()), dr["InvoiceNumber"].ToString(), "Error");
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > Null </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > " + Convert.ToString(dr["InvoiceNumber"]) + " </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["VendorName"])+" </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["Vendor"])+" </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["InvoiceCompany"])+" </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["InvoiceDate"])+" </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+Convert.ToString(dr["InvoiceAmount"])+" </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > Error </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "<td  > "+ ex.Message +" </td>";
                        BrokenInvoicesLog = BrokenInvoicesLog + "</tr>";
                    }

                }
            } else {
                Logger.log.Info("Header List doesnt exist or null");
            }
        }

        private static string PushtoDetailsList(DataRow dr, List listDetails, string sourceItemId)
        {
            string detailInserted = "False";
            Logger.log.Info("Pushing Invoice " + Convert.ToString(dr["InvoiceNumber"]) + " to details list with Header Source Item ID " + sourceItemId);
            Console.WriteLine("Processing Item " + sourceItemId);
            try
            {
                if (listDetails != null) {
                    string[] delimiter = { };
                    string[] companyArray = { };
                    string[] divisionArray = { };
                    string[] accountArray = { };
                    string[] DepartmentArray = { };
                    string[] AccountArray = { };
                    string[] JobArray = { };
                    string[] AmountArray = { };
                    string[] AmountDateArray = { };
                    string[] CERArray = { };
                    if (!string.IsNullOrEmpty(Convert.ToString(dr["Company"]))) {
                        if (dr["Account5CH"] != System.DBNull.Value && dr["ExpenseAmt"] != System.DBNull.Value) {
                            if (dr["Company"].ToString().Contains("\n") || dr["Company"].ToString().Contains(" ")) {
                                if (dr["Company"].ToString().Contains("\n")) {
                                    delimiter = new string[] { "\n" };
                                } else {
                                    delimiter = new string[] { " " };
                                }
                                companyArray = Convert.ToString(dr["Company"]).Split(delimiter, StringSplitOptions.None);
                                divisionArray = Convert.ToString(dr["Division"]).Split(delimiter, StringSplitOptions.None);
                                accountArray = Convert.ToString(dr["AccountString"]).Split(delimiter, StringSplitOptions.None);
                                DepartmentArray = Convert.ToString(dr["Dept3CH"]).Split(delimiter, StringSplitOptions.None); //malli
                                AccountArray = Convert.ToString(dr["Account5CH"]).Split(delimiter, StringSplitOptions.None);
                                JobArray = Convert.ToString(dr["Job"]).Split(delimiter, StringSplitOptions.None);
                                AmountArray = Convert.ToString(dr["ExpenseAmt"]).Split(delimiter, StringSplitOptions.None);
                                AmountDateArray = Convert.ToString(dr["AccntDate"]).Split(delimiter, StringSplitOptions.None);
                                CERArray = Convert.ToString(dr["CER"]).Split(delimiter, StringSplitOptions.None);

                            } else {
                                companyArray = new string[] { Convert.ToString(dr["Company"]) };
                                divisionArray = new string[] { Convert.ToString(dr["Division"]) };
                                accountArray = new string[] { Convert.ToString(dr["AccountString"]) };
                                DepartmentArray = new string[] { Convert.ToString(dr["Dept3CH"]) };
                                AccountArray = new string[] { Convert.ToString(dr["Account5CH"]) };
                                JobArray = new string[] { Convert.ToString(dr["Job"]) };
                                AmountArray = new string[] { Convert.ToString(dr["ExpenseAmt"]) };
                                AmountDateArray = new string[] { Convert.ToString(dr["AccntDate"]) };
                                CERArray = new string[] { Convert.ToString(dr["CER"]) };
                            }
                            for (int i = 0; i < companyArray.Length; i++) {
                                Console.WriteLine("Processing Array Item  " + i);
                                if (AccountArray.Length > i && !string.IsNullOrEmpty(Convert.ToString(AccountArray[i]))) {
                                    if (AmountArray.Length > i && !string.IsNullOrEmpty(Convert.ToString(AmountArray[i]))) {
                                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                        ListItem newItem = listDetails.AddItem(itemCreateInfo);
                                        newItem["Title"] = Convert.ToString(dr["InvoiceNumber"]);
                                        newItem["Invoice_x0020_Amount"] = Convert.ToString(dr["InvoiceAmount"]); //malli added{
                                        newItem["Invoice_x0020_Date"] = Convert.ToString(dr["InvoiceDate"]);
                                        newItem["PO_x0020_Number"] = Convert.ToString(dr["PO"]);
                                        newItem["Invoice_x0020_Company"] = Convert.ToString(dr["InvoiceCompany"]);
                                        newItem["Vendor_x0020_Number"] = Convert.ToString(dr["Vendor"]); //malli added}

                                        if (companyArray.Length > i)
                                            newItem["Company"] = Convert.ToString(companyArray[i]);
                                        else
                                            newItem["Company"] = string.Empty;
                                        // dr["Company"].ToString();
                                        if (divisionArray.Length > i) {
                                            newItem["Division"] = Convert.ToString(divisionArray[i]);//dr["Division"].ToString();
                                        } else {
                                            newItem["Division"] = string.Empty;
                                        }
                                        if (DepartmentArray.Length > i) {
                                            newItem["Department"] = Convert.ToString(DepartmentArray[i]);   //malli
                                        } else {
                                            newItem["Department"] = string.Empty;
                                        }
                                        newItem["Account"] = Convert.ToString(AccountArray[i]);
                                        newItem["Amount"] = Convert.ToString(AmountArray[i]);
                                        newItem["Job_x0020__x0023_"] = "";

                                        if (JobArray.Length > i) {
                                            newItem["Job_x0020__x0023_"] = Convert.ToString(JobArray[i]);
                                        } else {
                                            newItem["Job_x0020__x0023_"] = string.Empty;
                                        }

                                        if (CERArray.Length > i) {
                                            newItem["CER_x0020__x0023_"] = Convert.ToString(CERArray[i]);
                                        } else {
                                            newItem["CER_x0020__x0023_"] = string.Empty;
                                        }
                                        newItem["Account_x0020_Date"] = "";
                                        if (AmountDateArray.Length > i) {
                                            newItem["Account_x0020_Date"] = Convert.ToString(AmountDateArray[i]);//from excel
                                        } else {
                                            newItem["Account_x0020_Date"] = string.Empty;
                                        }

                                        newItem["Account_x0020_String"] = "";
                                        if (accountArray.Length > i) {
                                            newItem["Account_x0020_String"] = Convert.ToString(accountArray[i]);//dr["AccountString"].ToString();
                                        } else {
                                            newItem["Account_x0020_String"] = string.Empty;
                                        }

                                        newItem["SourceItemId"] = sourceItemId;
                                        newItem.Update();

                                        if (AmountArray.Length > i && AccountArray.Length > i) {
                                            Logger.log.Info("Processed the Details Item with the invoice Number :" + Convert.ToString(dr["InvoiceNumber"]) + " Amount : " + Convert.ToString(AmountArray[i]) + " Account : " + Convert.ToString(AccountArray[i]));
                                        }
                                        Logger.log.Info("Processed the Details Item with the Account Date :" + Convert.ToString(newItem["Account_x0020_Date"]) + " AccountString : " + Convert.ToString(newItem["Account_x0020_String"]) + " Job_No : " + Convert.ToString(newItem["Job_x0020__x0023_"]));
                                        if (companyArray.Length > i) {
                                            Logger.log.Info("Processed the Details Item with the Company :" + Convert.ToString(companyArray[i]) + " Vendor Number : " + Convert.ToString(dr["Vendor"]) + " Invoice Company : " + Convert.ToString(dr["InvoiceCompany"]));
                                        }
                                        if (divisionArray.Length > i && DepartmentArray.Length > i) {
                                            Logger.log.Info("Processed the Details Item with the Division :" + Convert.ToString(divisionArray[i]) + " Department : " + Convert.ToString(DepartmentArray[i]) + " Invoice Date : " + Convert.ToString(dr["InvoiceDate"]));
                                        }
                                        Logger.log.Info("Processed the Details Item with the Title :" + Convert.ToString(dr["InvoiceNumber"]));
                                    }
                                }
                            }
                            Logger.log.Info("Execute Query");
                            context.ExecuteQuery();
                            detailInserted = "True";
                            processedCount = processedCount + 1;
                            Logger.log.Info("Completed Push to Details List for Invoice " + Convert.ToString(dr["InvoiceNumber"]));
                        } else {
                            Logger.log.Info(" The Account or Expense Amount is null for the Invoice " + Convert.ToString(dr["InvoiceNumber"]) + " to details list with Header Source Item ID " + sourceItemId);
                        }
                    } else {
                        Logger.log.Info(" The company values is null for the Invoice " + Convert.ToString(dr["InvoiceNumber"]) + " to details list with Header Source Item ID " + sourceItemId);
                    }
                }
            } catch (Exception ex) {
                detailInserted = "Error";
                Logger.log.Info("Error in pushing details list for Invoice  " + Convert.ToString(dr["InvoiceNumber"]) + "</br>" + "Error " + ex.Message);
                // BrokenInvoicesLog = BrokenInvoicesLog + "<br /><br />" + " Error in Pushing Invoice to Details, Invoice: " + Convert.ToString(dr["InvoiceNumber"]) + ", Error: " + ex.Message;
                ErrorMessage = ex.Message;

            }
            return detailInserted;
        }
        private static void UpdateCSInvoiceStatus(int coversheetItemID, string InvoiceNumber, string result)
        {
            try
            {
                Logger.log.Info("Updating Coversheet Invoice : " + InvoiceNumber + " With Item ID : " + coversheetItemID + " with Status : " + result);
                if (listCoversheet != null) {
                    ListItem coverSheetInvoice = listCoversheet.GetItemById(coversheetItemID);

                    if (coverSheetInvoice != null) {
                        coverSheetInvoice["processed"] = result;
                        coverSheetInvoice.Update();
                        context.Load(coverSheetInvoice);
                        context.ExecuteQuery();
                    }
                    Logger.log.Info("Updated Coversheet Invoice " + InvoiceNumber + " with Status " + result);
                }
            } catch (Exception ex) {
                Logger.log.Info("Error in updating Coversheet Invoice Status " + ex.Message);
                BrokenInvoicesLog = BrokenInvoicesLog + "<br /><br />" + "Error in Updating Coversheet Invoice " + InvoiceNumber + " recordd's Status : " + ex.Message;
            }
        }

        public static string GetFileFromFolderPath(string targetDirectory)
        {

            Logger.log.Info("Getting file from Folder :  " + targetDirectory);
            string filePath = string.Empty;
            string[] fileEntries = null;
            try
            {
                fileEntries = Directory.GetFiles(targetDirectory);
                if (fileEntries != null) {
                    if (fileEntries.Length > 0) {
                        filePath = fileEntries[0];

                    }
                }
            } catch (Exception ex) {
                Logger.log.Info(ex.Message);
            }
            return filePath;
        }

        private static void MoveFile(string source, string destination)
        {
            try
            {
                Logger.log.Info("Moving Voucher Mappings file to destination " + destination);
                System.IO.File.Move(source, destination);
            }
            catch (Exception ex)
            {
                Logger.log.Warn("Error in moving the Voucher Mappings file to destination :" + destination);
                BrokenInvoicesLog = BrokenInvoicesLog + "<br /><br />" + "Error in moving the Voucher Mappings file to destination " + "Error " + ex.Message;
            }
        }
 
        private static Boolean MapVoucher(DataTable dtVouchers, List listHeader)
        {
            Logger.log.Info("Mapping Process Started..");
            bool mappingDone = false;
            try
            {
                int mappingCounter = 0;
                if (listHeader != null)
                {
                    Logger.log.Info("Mapping file contains " + dtVouchers.Rows.Count + " entries ");
                    foreach (DataRow dr in dtVouchers.Rows)
                    {
                        Logger.log.Info("Processing mapping for vendor Invoice number : " + dr["Vendor invoice number"].ToString().Trim());
                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                            "<Value Type='Text'>" + dr["Vendor invoice number"].ToString().Trim() + " </Value></Eq></Where></Query></View>";
                        ListItemCollection collection = listHeader.GetItems(camlQuery);
                        context.Load(collection, items => items.Include(
                                    item => item.Id, item => item["Title"],
                                    item => item["Vendor_x0020_Name"],
                                    item => item["Vendor_x0020_Number"],
                                    item => item["Invoice_x0020_Amount"],
                                    item => item["Voucher_x0020_Date"],
                                    item => item["Payment_x0020_Status"],
                                    item => item["Accounting_x0020_Approval"]
                                    ));
                        context.ExecuteQuery();
                        if (collection.Count > 0)
                        {
                            Logger.log.Info("Found " + collection.Count + " entriesin the header " + dr["Vendor invoice number"].ToString().Trim());
                            foreach (ListItem item in collection)
                            {
                                if (item["Accounting_x0020_Approval"].ToString() == "Accounting Approval Approved")
                                {
                                    if (item["Vendor_x0020_Number"] != null && item["Vendor_x0020_Number"].ToString().Equals(dr["Vendor ID"].ToString().Trim(), StringComparison.OrdinalIgnoreCase))
                                    {
                                        if (item["Invoice_x0020_Amount"] != null && item["Invoice_x0020_Amount"].ToString().Equals(dr["Gross amt relieved ICUR"].ToString().Trim(), StringComparison.OrdinalIgnoreCase))
                                        {
                                            item["Voucher_x0020_Number"] = dr["Payment Reference"].ToString();
                                            item["Voucher_x0020_Date"] = dr[" Payment Date Filter"];// do +1
                                            item["Payment_x0020_Status"] = "Paid";
                                            item.Update();
                                            context.Load(item);
                                            context.ExecuteQuery();
                                            Logger.log.Info("Mapping is completed for item with Voucher Number " + item["Voucher_x0020_Number"].ToString() + " Vendor Number " + item["Vendor_x0020_Number"].ToString() + "Vendor ID " + dr["Vendor ID"].ToString());
                                            mappingCounter++;
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Logger.log.Info("No entries found in the header with Invoice Number : " + dr["Vendor invoice number"].ToString());
                        }
                    }
                    if (mappingCounter > 0)
                        mappingDone = true;
                    else
                        mappingDone = false;
                }
            }
            catch (Exception ex)
            {
                mappingDone = false;
                Console.WriteLine(" Error {0}", ex);
                Logger.log.Info("Error in mapping the vouchers " + ex.Message);
                BrokenInvoicesLog = BrokenInvoicesLog + "<br /><br />" + "Error in mapping the vouchers " + "Error " + ex.Message;
            }

            return mappingDone;
        }

        private static DataTable ConstructDataTable()
       {
            Logger.log.Info("Constructing Data Table for Invoices");
            dtInvoices = new DataTable();
            dtInvoices.Columns.Add("ID");
            dtInvoices.Columns.Add("VendorName");
            dtInvoices.Columns.Add("Vendor");
            dtInvoices.Columns.Add("PO");
            dtInvoices.Columns.Add("InvoiceNumber");
            dtInvoices.Columns.Add("InvoiceDate");
            dtInvoices.Columns.Add("InvoiceCompany");
            dtInvoices.Columns.Add("InvoiceDivision");
            dtInvoices.Columns.Add("InvoiceAmount");
            dtInvoices.Columns.Add("Company");
            dtInvoices.Columns.Add("Division");
            dtInvoices.Columns.Add("Dept3CH");
            dtInvoices.Columns.Add("Account5CH");
            dtInvoices.Columns.Add("ExpenseAmt");
            dtInvoices.Columns.Add("Job");
            dtInvoices.Columns.Add("CER");
            // dtInvoices.Columns.Add("MonetaryUnit");
            dtInvoices.Columns.Add("AccntDate");
            dtInvoices.Columns.Add("AccountString");
            //add another column here
            return dtInvoices;

        }
 
        private static bool CheckIfInvoiceAlreadyProcessed(List listDetails, DataRow dr)
       {
            bool exist = false;
            try
            {
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                    "<Value Type='Text'>" + dr["InvoiceNumber"].ToString().Trim() + " </Value></Eq></Where></Query></View>";
                ListItemCollection collection = listDetails.GetItems(camlQuery);
                context.Load(collection, items => items.Include(
                          item => item.Id, item => item["Title"]
                         ));
                context.ExecuteQuery();
                if (collection != null)
                {
                    if (collection.Count > 0)
                    {
                        exist = true;
                        Console.WriteLine("The Invoice : " + Convert.ToString(dr["InvoiceNumber"]).Trim() + " already processed, Hence Ignoring");
                    }
                }
            }
            catch (Exception ex)
            {
                //BrokenInvoices Email Body
                BrokenInvoicesLog = BrokenInvoicesLog + "<br /><br />" + " Error in Processing Invoice " + Convert.ToString(dr["InvoiceNumber"]) + " Error : " + ex.Message;
            }

            return exist;
        }

        //Method to get Distinct Invoice Ids from Header and store in Array List
        private static ArrayList GetHeaderInvoiceIds(List listHeader)
        {
            Logger.log.Info("Getting Distinct Invoices Ids from Header");
            ArrayList arrayInvoiceHeader = new ArrayList();
            try
            {
                CamlQuery camlQueryHeader = new CamlQuery();
                camlQueryHeader.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                                                    "<Value Type='Number'>0</Value></Geq></Where></Query></View>";
                ListItemCollection collectionHeader = listHeader.GetItems(camlQueryHeader);
                context.Load(collectionHeader,
                             items => items.Include(item => item["Title"] //Invoice No#
                            ));
                context.ExecuteQuery();
                foreach (ListItem item in collectionHeader)
                {
                    if (item["Title"] != null)
                    {
                        if (!arrayInvoiceHeader.Contains(item["Title"].ToString()))
                        {
                            arrayInvoiceHeader.Add(item["Title"].ToString());
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.log.Error("Error in Retrieving Invoice Ids from Header " + ex.Message);
                BrokenInvoicesLog = BrokenInvoicesLog + "<br /><br />" + "Error in Retrieving Invoice Ids from Header " + ex.Message;
            }

            return arrayInvoiceHeader;
        }

    }
}
