using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using log4net;


namespace Dayton_RogersTool
{
    class Utils
    {
        public static void ReadConfigurationValues()
        {
            try
            {
                Console.WriteLine("Reading Config");
                Logger.log.Info("Reading Configurations from AppConfig..");
                Invoice._invoiceSourceLibraryName = ConfigurationManager.AppSettings.Get("CoversheetList");
                Invoice._invoceDestinationListName = ConfigurationManager.AppSettings.Get("CoversheetHeaderList");
                Invoice._coversheetDetailsListName = ConfigurationManager.AppSettings.Get("CoversheetDetailsList");
                Invoice._siteUrl = ConfigurationManager.AppSettings.Get("SiteUrl");
                Invoice._domain = ConfigurationManager.AppSettings.Get("Domain");
                Invoice._userName = ConfigurationManager.AppSettings.Get("UserName");
                Invoice._password = ConfigurationManager.AppSettings.Get("Password");
                // Invoice._cheqRegListName = ConfigurationManager.AppSettings.Get("ChequeList");//Cheque Register : New Item To be Pushed
                // Invoice._vendorInvoiceListName = ConfigurationManager.AppSettings.Get("VendorList");//VendorInvoiceList :          
                Invoice._filePath = ConfigurationManager.AppSettings.Get("InputFilePath");
                Invoice._repository = ConfigurationManager.AppSettings.Get("Repository");
                Invoice._mapVouchers = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("MapVoucher"));
                Invoice._runforCloud = Convert.ToBoolean(ConfigurationManager.AppSettings.Get("CloudEnvironment"));

            }
            catch (Exception ex)
            {
                Logger.log.Error("Error in Reading Configurations from AppConfig.."+ex.Message);
                throw ex;
            }
        }
        public static DataTable ReadExcelData(string filePath)
        {
            DataTable dtResult = new DataTable();
            string fileExtension = filePath.Split('.')[1].ToString();
            OleDbConnection objConn = new OleDbConnection();
            try
            {
                //string sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + CurrentFilePath + ";Extended Properties=Excel 8.0;";
                string sConnectionString = "";
                if (fileExtension.ToLower() == "xls")
                {
                    sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1;\"";
                }
                else if (fileExtension.ToLower() == "xlsx")
                {
                    sConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1;\"";
                }
                objConn = new OleDbConnection(sConnectionString);
                objConn.Open();
                DataTable dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM [" + excelSheets[0].ToString() + "]", objConn);
                OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();
                objAdapter1.SelectCommand = objCmdSelect;
                DataSet dsData = new DataSet();
                objAdapter1.Fill(dsData, "XLData");
                objConn.Close();
                for (int cnt = 0; cnt < dsData.Tables[0].Rows.Count; cnt++)
                {
                    bool flag = true;
                    for (int colcnt = 0; colcnt < dsData.Tables[0].Columns.Count; colcnt++)
                    {
                        if (dsData.Tables[0].Rows[cnt][dsData.Tables[0].Columns[colcnt].ToString()].ToString().Trim() != "")
                        {
                            flag = false;
                            break;
                        }
                    }
                    if (flag)
                    {
                        dsData.Tables[0].Rows.RemoveAt(cnt);
                        cnt = cnt - 1;
                    }
                }
                for (int col = 0; col < dsData.Tables[0].Columns.Count; col++)
                {
                    string temp = "";
                    string temp1 = "";
                    string temp2 = "";
                    temp = dsData.Tables[0].Columns[col].ToString().ToLower();
                    if (temp.Length > 2)
                    {
                        temp1 = temp.Substring(0, 1);
                        temp2 = temp.Substring(1, 1);
                        if (temp1 == "f" && temp2 == "0" || temp2 == "1" || temp2 == "2" || temp2 == "3" || temp2 == "4" || temp2 == "5" || temp2 == "6" || temp2 == "7" || temp2 == "8" || temp2 == "9")
                        {
                            dsData.Tables[0].Columns.RemoveAt(col);
                            col = col - 1;
                        }
                    }
                    else
                    {
                        dsData.Tables[0].Columns.RemoveAt(col);
                        col = col - 1;
                    }
                    dsData.Tables[0].AcceptChanges();
                    dtResult = dsData.Tables[0];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (objConn != null)
                {
                    objConn.Close();
                }
            }
            return dtResult;
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
