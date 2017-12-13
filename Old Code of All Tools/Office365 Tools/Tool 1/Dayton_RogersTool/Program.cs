using System;
using Microsoft.SharePoint.Client;
using System.Data;
using System.Net;
using System.IO;
using log4net;
using System.Diagnostics;


namespace Dayton_RogersTool
{
    class Program
    {
        #region variables
        static ClientContext context;
        static DataTable dtInvoices;
        static DataTable dtExcel;
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
                Utils.ReadConfigurationValues();
                dtInvoices = ConstructDataTable();
                if (!string.IsNullOrEmpty(Invoice._siteUrl))
                {
                    context = new ClientContext(Invoice._siteUrl);
                }
                else
                {
                    Logger.log.Info("The site url is empty in configuration..");
                }

                if (Invoice._runforCloud)
                {
                    Logger.log.Info("Configuration set to run for Cloud Environment..");
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
                    Logger.log.Info("Configuration set to run for On Premise Environment..");
                    context.Credentials = new NetworkCredential(Invoice._userName, Invoice._password, Invoice._domain);
                }

                Web web = context.Web;
                context.ExecuteQuery();
                Logger.log.Info("Context created for given site " + Invoice._siteUrl);

                List listCoversheet = web.Lists.GetByTitle(Invoice._invoiceSourceLibraryName);
                List listHeader = web.Lists.GetByTitle(Invoice._invoceDestinationListName);
                List listDetails = web.Lists.GetByTitle(Invoice._coversheetDetailsListName);//To push
                Logger.log.Info("Reading items from List : " + Invoice._invoiceSourceLibraryName);
                Logger.log.Info("Reading items from List : " + Invoice._invoiceSourceLibraryName);
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View><Query><Where><Geq><FieldRef Name='ID'/>" +
                                                    "<Value Type='Number'>0</Value></Geq></Where></Query></View>";
                ListItemCollection collection = listCoversheet.GetItems(camlQuery);
                Logger.log.Info("Reading items from List : " + camlQuery.ViewXml);
                context.Load(collection, items => items.Include(
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
                                                         //      item => item["Monetary_Unit"],
                                                         item => item["Acct_x0020_Date"],
                                                         item => item["Account_x0020_String"]
                                                         ));
                context.ExecuteQuery();
                dtInvoices.Clear();
                Logger.log.Info("Filling datatable with the items of List : " + Invoice._invoiceSourceLibraryName);
                foreach (ListItem item in collection)
                {
                    DataRow dr = dtInvoices.NewRow();
                    dr["ID"] = item.Id.ToString();
                    if (item["Vendor_x0020_Name"] != null)
                        dr["VendorName"] = item["Vendor_x0020_Name"].ToString();
                    if (item["Vendor"] != null)
                        dr["Vendor"] = item["Vendor"].ToString();
                    if (item["PO"] != null)
                        dr["PO"] = item["PO"].ToString();
                    if (item["Invoice_x0020__x0023_"] != null)
                        dr["InvoiceNumber"] = item["Invoice_x0020__x0023_"].ToString();
                    if (item["Invoice_x0020_Date"] != null)
                        dr["InvoiceDate"] = item["Invoice_x0020_Date"].ToString();
                    if (item["Invoice_x0020_Co"] != null)
                        dr["InvoiceCompany"] = item["Invoice_x0020_Co"].ToString();
                    if (item["Invoice_x0020_DIV"] != null)
                        dr["InvoiceDivision"] = item["Invoice_x0020_DIV"].ToString();
                    if (item["Invoice_x0020_Amt"] != null)
                        dr["InvoiceAmount"] = item["Invoice_x0020_Amt"].ToString();
                    if (item["Company"] != null)
                        dr["Company"] = item["Company"].ToString();
                    if (item["Division"] != null)
                        dr["Division"] = item["Division"].ToString();
                    if (item["Dept_x002d_3_x0020_CH"] != null)
                        dr["Dept3CH"] = item["Dept_x002d_3_x0020_CH"].ToString();
                    if (item["Account_x002d_5_x0020_CH"] != null)
                        dr["Account5CH"] = item["Account_x002d_5_x0020_CH"].ToString();
                    if (item["Expense_x0020_Amt"] != null)
                        dr["ExpenseAmt"] = item["Expense_x0020_Amt"].ToString();
                    if (item["Job"] != null)
                        dr["Job"] = item["Job"].ToString();
                    if (item["CER"] != null)
                        dr["CER"] = item["CER"].ToString();
                    // if (item["Monetary_Unit"] != null)
                    //     dr["MonetaryUnit"] = item["Monetary_Unit"].ToString();
                    if (item["Acct_x0020_Date"] != null)
                        dr["AccntDate"] = item["Acct_x0020_Date"].ToString();
                    if (item["Account_x0020_String"] != null)
                        dr["AccountString"] = item["Account_x0020_String"].ToString();
                    dtInvoices.Rows.Add(dr);
                }
                if (dtInvoices != null && dtInvoices.Rows.Count > 0)
                {
                    Logger.log.Info("Retrieved " + dtInvoices.Rows.Count + " Invoices from CoverSheet ");
                    PushtoHeaderList(dtInvoices, listHeader, listDetails);
                }
               Console.WriteLine("Invoice Processing Completed!");
               Console.WriteLine("Splitting is Completed.");
                Logger.log.Info("Splitting Process is Completed.. ");
                if (Invoice._mapVouchers)
                {
                    Logger.log.Info("Starting Mapping the Voucher Numbers");
                   Console.WriteLine("Started Mapping Voucher Number!");
                    //Console.WriteLine("Verifly the invoice and approve, Again run the application...");
                    if (!string.IsNullOrEmpty(Invoice._filePath))
                    {
                        string filepath = string.Empty;
                        if (System.IO.File.Exists(Invoice._filePath))
                        {
                            Logger.log.Info("Getting voucher mappings file from filepath");
                            filepath = Invoice._filePath;
                        }
                        else
                        {
                            Logger.log.Info("Getting voucher mappings file from folder path");
                           Console.WriteLine("Getting file from folder  :" + Invoice._filePath);
                            filepath = GetFileFromFolderPath(Invoice._filePath);
                        }
                        if (!string.IsNullOrEmpty(filepath))
                        {

                            string destinationPath = string.Empty;
                            if (System.IO.File.Exists(Invoice._repository))
                            {
                                Logger.log.Info("Reading destination mapping vouchers filepath ");
                                destinationPath = Invoice._repository;
                            }
                            else
                            {
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
                            dtVouchers = Utils.ReadExcelData(filepath);
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
                            }
                        }
                        else
                        {
                           Console.WriteLine("Unable to map the voucher, reason could be the file path is not correct or the file doenst exist in the specified locaton : " + filepath);
                            Logger.log.Info("Unable to map the voucher, reason could be the file path is not correct or the file doenst exist in the specified locaton : " + filepath);
                        }
                    }
                }
                else
                {
                    Logger.log.Info("Mapping Voucher set to false in configuration.");
                }
                Logger.log.Info("============= Completed Invoice Processing ====================");
               Console.WriteLine("Please Close the Window to exit Process..!");
         //       Console.ReadLine();

                #endregion InvoiceSplit

            }
            catch (Exception ex)
            {
                //Console.ForegroundColor = ConsoleColor.Red;
               Console.WriteLine("Error in Processing Invoice, Please check all the configurations and try again..{0}", ex);
                Logger.log.Error("Error in Processing Invoices " + ex.Message);
               // Console.ResetColor();
               // Console.ReadLine();
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
                if (fileEntries != null)
                {
                    if (fileEntries.Length > 0)
                    {
                        filePath = fileEntries[0];

                    }
                }
            }
            catch (Exception ex)
            {
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
            catch (Exception)
            {
                Logger.log.Warn("Error in moving the Voucher Mappings file to destination :" + destination);
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
            }

            return mappingDone;
        }
        private static void PushtoHeaderList(DataTable dtInvoices, List listHeader, List listDetails)
        {
            Logger.log.Info("Pushing Invoice data to header list");
            if (listHeader != null)
            {
                foreach (DataRow dr in dtInvoices.Rows)
                {
                    if (dr["InvoiceNumber"] != null) //malli added dr["Account5CH"] != DBNull.Value
                    {
                        Logger.log.Info("Processing Invoice " + dr["InvoiceNumber"].ToString());
                        bool InvoiceExist = CheckIfInvoiceAlreadyProcessed(listHeader, dr);
                        if (!InvoiceExist)
                        {
                            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                            ListItem newItem = listHeader.AddItem(itemCreateInfo);
                            newItem["Title"] = dr["InvoiceNumber"].ToString();//vendor invoice number
                            newItem["Vendor_x0020_Name"] = dr["VendorName"].ToString();
                            newItem["Vendor_x0020_Number"] = dr["Vendor"].ToString();
                            newItem["PO_x0020_Number"] = dr["PO"].ToString();
                            newItem["Invoice_x0020_Company"] = dr["InvoiceCompany"].ToString();
                            newItem["Invoice_x0020_Division"] = dr["InvoiceDivision"].ToString();
                            newItem["Invoice_x0020_Date"] =dr["InvoiceDate"].ToString();
                            newItem["Invoice_x0020_Amount"] = dr["InvoiceAmount"].ToString();
                            newItem.Update();
                            context.ExecuteQuery();
                            Logger.log.Info("Invoice " + dr["InvoiceNumber"].ToString() + "Pushed to Header List");
                            string sourceItemId = newItem.Id.ToString();
                            PushtoDetailsList(dr, listDetails, sourceItemId);
                        }
                        else
                        {
                            Logger.log.Info("Invoice " + dr["InvoiceNumber"].ToString() + " already exist in Header List");
                        }
                    }
                }
            }
        }
        private static bool CheckIfInvoiceAlreadyProcessed(List listHeader, DataRow dr)
        {
            bool exist = false;
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Title'/>" +
                "<Value Type='Text'>" + dr["InvoiceNumber"].ToString().Trim() + " </Value></Eq></Where></Query></View>";
            ListItemCollection collection = listHeader.GetItems(camlQuery);
            context.Load(collection, items => items.Include(
                      item => item.Id, item => item["Title"]
                     ));
            context.ExecuteQuery();

            if (collection != null)
            {
                if (collection.Count > 0)
                {
                    exist = true;
                    Console.WriteLine("The Invoice : " + dr["InvoiceNumber"].ToString().Trim() + " already processed, Hence Ignoring");
                }
            }
            return exist;
        }

        private static void PushtoDetailsList(DataRow dr, List listDetails, string sourceItemId)
        {
            Logger.log.Info(" Pushing Invoice " + dr["InvoiceNumber"].ToString() + " to details list with Header Source Item ID " + sourceItemId);
           Console.WriteLine("Processing Item " + sourceItemId);
            if (listDetails != null)
            {
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

                if (dr["Company"] != null)
                {
                    if (dr["Account5CH"] != System.DBNull.Value && dr["ExpenseAmt"] != System.DBNull.Value)
                    {
                        if (dr["Company"].ToString().Contains("\n") || dr["Company"].ToString().Contains(" ")) //malli
                        {
                            if (dr["Company"].ToString().Contains("\n"))
                            {
                                delimiter = new string[] { "\n" };
                            }
                            else
                            {
                                delimiter = new string[] { " " };
                            }
                            companyArray = dr["Company"].ToString().Split(delimiter, StringSplitOptions.None);
                            divisionArray = dr["Division"].ToString().Split(delimiter, StringSplitOptions.None);
                            accountArray = dr["AccountString"].ToString().Split(delimiter, StringSplitOptions.None);
                            DepartmentArray = dr["Dept3CH"].ToString().Split(delimiter, StringSplitOptions.None); //malli
                            AccountArray = dr["Account5CH"].ToString().Split(delimiter, StringSplitOptions.None);
                            JobArray = dr["Job"].ToString().Split(delimiter, StringSplitOptions.None);
                            AmountArray = dr["ExpenseAmt"].ToString().Split(delimiter, StringSplitOptions.None);
                            AmountDateArray = dr["AccntDate"].ToString().Split(delimiter, StringSplitOptions.None);
                            CERArray = dr["CER"].ToString().Split(delimiter, StringSplitOptions.None);

                        }
                        else
                        {
                            companyArray = new string[] { dr["Company"].ToString() };
                            divisionArray = new string[] { dr["Division"].ToString() };
                            accountArray = new string[] { dr["AccountString"].ToString() };
                            DepartmentArray = new string[] { dr["Dept3CH"].ToString() };
                            AccountArray = new string[] { dr["Account5CH"].ToString() };
                            JobArray = new string[] { dr["Job"].ToString() };
                            AmountArray = new string[] { dr["ExpenseAmt"].ToString() };
                            AmountDateArray = new string[] { dr["AccntDate"].ToString() };
                            CERArray = new string[] { dr["CER"].ToString() };
                        }
                        for (int i = 0; i < companyArray.Length; i++)
                        {
                           Console.WriteLine("Processing Array Item  " + i);
                            if (AccountArray.Length > i && AccountArray[i] != null)
                            {
                                if (AmountArray.Length > i && AmountArray[i] != null)
                                {
                                    ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                                    ListItem newItem = listDetails.AddItem(itemCreateInfo);
                                    newItem["Title"] = dr["InvoiceNumber"].ToString();
                                    newItem["Invoice_x0020_Amount"] = dr["InvoiceAmount"].ToString(); //malli added{
                                    newItem["Invoice_x0020_Date"] = dr["InvoiceDate"].ToString();
                                    newItem["PO_x0020_Number"] = dr["PO"].ToString();
                                    newItem["Invoice_x0020_Company"] = dr["InvoiceCompany"].ToString();
                                    newItem["Vendor_x0020_Number"] = dr["Vendor"].ToString(); //malli added}
                                    newItem["Company"] = companyArray[i].ToString();// dr["Company"].ToString();
                                    newItem["Division"] = divisionArray[i].ToString();//dr["Division"].ToString();
                                    newItem["Department"] = DepartmentArray[i].ToString();   //malli
                                    newItem["Account"] = AccountArray[i].ToString();
                                    newItem["Amount"] = AmountArray[i].ToString();
                                    newItem["Job_x0020__x0023_"] = "";
                                    if (JobArray.Length>i)
                                    { newItem["Job_x0020__x0023_"] = JobArray[i].ToString(); }
                                    if (CERArray.Length > i)
                                    { newItem["CER_x0020__x0023_"] = CERArray[i].ToString(); }
                                    newItem["Account_x0020_Date"] = "";
                                    if (AmountDateArray.Length > i)
                                    {
                                        newItem["Account_x0020_Date"] = AmountDateArray[i].ToString();//from excel
                                    }
                                    newItem["Account_x0020_String"] = "";
                                    if (accountArray.Length > i)
                                    {
                                        newItem["Account_x0020_String"] = accountArray[i].ToString();//dr["AccountString"].ToString();
                                    }


                                    newItem["SourceItemId"] = sourceItemId;
                                    newItem.Update();
                                    Logger.log.Info("Processed the Details Item with the invoice Number :" + dr["InvoiceNumber"].ToString() + " Amount : " + AmountArray[i].ToString() + " Account : " + AccountArray[i].ToString());
                                    Logger.log.Info("Processed the Details Item with the Account Date :" + newItem["Account_x0020_Date"] + " AccountString : " + newItem["Account_x0020_String"] + " Job_No : " + newItem["Job_x0020__x0023_"]);
                                    Logger.log.Info("Processed the Details Item with the Company :" + companyArray[i].ToString() + " Vendor Number : " + dr["Vendor"].ToString() + " Invoice Company : " + dr["InvoiceCompany"].ToString());
                                    Logger.log.Info("Processed the Details Item with the Division :" + divisionArray[i].ToString() + " Department : " + DepartmentArray[i].ToString() + " Invoice Date : " + dr["InvoiceDate"].ToString());
                                    Logger.log.Info("Processed the Details Item with the Title :" + dr["InvoiceNumber"].ToString());

                                }
                            }
                        }
                        Logger.log.Info("Execute Query");
                        context.ExecuteQuery();
                        Logger.log.Info("Completed Push to Details List for Invoice " + dr["InvoiceNumber"].ToString());

                    }
                    else
                    {
                        Logger.log.Info(" The Account or Expense Amount is null for the invoice " + dr["InvoiceNumber"].ToString() + " to details list with Header Source Item ID " + sourceItemId);
                    }
                }
                else
                {
                    Logger.log.Info(" The company values is null for the invoice " + dr["InvoiceNumber"].ToString() + " to details list with Header Source Item ID " + sourceItemId);
                }
            }

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
    }
}
