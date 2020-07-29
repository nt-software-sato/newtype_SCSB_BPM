using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Configuration;
using NAWXDBCINFOIOLib;
using System.IO;
using System.Net;
using Microsoft.Win32;
using NewType.FlowSe7en.FlowEngine;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;


public partial class A66CallFERP : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string CASENO = HttpUtility.HtmlEncode(Request.QueryString["CASENO"]) ?? string.Empty;
        string CASETYPE = HttpUtility.HtmlEncode(Request.QueryString["CASETYPE"]) ?? string.Empty;
        string PAYMENTPERIOD = HttpUtility.HtmlEncode(Request.QueryString["PAYMENTPERIOD"]) ?? string.Empty;
        string ISCONVERTPREPAYMENT = HttpUtility.HtmlEncode(Request.QueryString["ISCONVERTPREPAYMENT"]) ?? string.Empty;
        string ALLOCATIONSTARTYEARMONTH = HttpUtility.HtmlEncode(Request.QueryString["ALLOCATIONSTARTYEARMONTH"]) ?? string.Empty;
        string ALLOCATIONENDYEARMONTH = HttpUtility.HtmlEncode(Request.QueryString["ALLOCATIONENDYEARMONTH"]) ?? string.Empty;
	double WITHHOLDINGINCOMETAXRATIO = Convert.ToDouble(HttpUtility.HtmlEncode(Request.QueryString["WITHHOLDINGINCOMETAXRATIO"]));
        string WITHHOLDINGINCOMEACCOUNTID = HttpUtility.HtmlEncode(Request.QueryString["WITHHOLDINGINCOMEACCOUNTID"]) ?? string.Empty;
        double WITHHOLDINGINHEALTHINSURANCETAXRATIO = Convert.ToDouble(HttpUtility.HtmlEncode(Request.QueryString["WITHHOLDINGINHEALTHINSURANCETAXRATIO"]));
        string WITHHOLDINGINHEALTHINSURANCEPREMIUMACCOUNTID = HttpUtility.HtmlEncode(Request.QueryString["WITHHOLDINGINHEALTHINSURANCEPREMIUMACCOUNTID"]) ?? string.Empty;
        string RequisitionID = HttpUtility.HtmlEncode(Request.QueryString["RequisitionID"]) ?? string.Empty;
        string Type = HttpUtility.HtmlEncode(Request.QueryString["Type"]) ?? string.Empty;
        string USECASE="";

        //宣告要ERP起單的表單名稱
        string JsonText = "";
        string CommandText = "";
        string FErpServer = "";
        string TmpSql = "";
        string ErrMsg = "";
        string DataText = "";
        DataTable TmpDt = new DataTable();
        DataTable dt = new DataTable();
        DataTable DetailDt = new DataTable();
        List<SqlParameter> ParameterList = new List<SqlParameter>();

        //RequisitionID = "2a9a48f0-aa1c-4d87-87bd-0eee6b730dfc";
        //使用AutoWeb資料庫連線字串
        string connString = LoadCmdStr("\\\\Database\\\\Project\\\\SCSB\\\\Flow\\\\connection\\\\BPM.xdbc.xmf", 1);

        //使用Sql連線並傳回Datatable
        /*
        Type = "1";
        CASENO = "20190407-0225";
        CASETYPE = "1";
        PAYMENTPERIOD = "1";
        ISCONVERTPREPAYMENT = "N";
        ALLOCATIONSTARTYEARMONTH = "201905";
        ALLOCATIONENDYEARMONTH = "201908";
        RequisitionID = "850efd78-03bd-41b9-afcf-ece8b2160ee1";
        */
        USECASE = "UC_EM_ENTRYWHOLEBATCH";
        
        string CallUrl = @"/LCJsonBridge/JsonService.do";
        //用Linq直接組Json

        if (Type == "2")
        {
            var Result = new
            {
                SENDERACCOUNT = "ADMIN",
                SENDERPASSWORD = "ADMIN0424",
                SENDERTENANTID = "0081",
                HOSTINFO = "NewType",
                HOSTLANGUAGE = "TraditionalChinese",
                CASENO = CASENO,
                CASETYPE = CASETYPE,
                PAYMENTPERIOD = PAYMENTPERIOD,
                ISCONVERTPREPAYMENT = ISCONVERTPREPAYMENT,
                ALLOCATIONSTARTYEARMONTH = ALLOCATIONSTARTYEARMONTH,
                ALLOCATIONENDYEARMONTH = ALLOCATIONENDYEARMONTH,
		WITHHOLDINGINCOMETAXRATIO = WITHHOLDINGINCOMETAXRATIO,
                WITHHOLDINGINCOMEACCOUNTID = WITHHOLDINGINCOMEACCOUNTID,
                WITHHOLDINGINHEALTHINSURANCETAXRATIO = WITHHOLDINGINHEALTHINSURANCETAXRATIO,
                WITHHOLDINGINHEALTHINSURANCEPREMIUMACCOUNTID = WITHHOLDINGINHEALTHINSURANCEPREMIUMACCOUNTID,
                NEWTYPEBPMREQUISITIONID = RequisitionID,
                USECASE = USECASE,
                SERVICEID = "EXEPURCHASEPAYMENTENTRYBYACREATE"
            };
            //序列化為JSON字串並輸出結果
            JsonText = JsonConvert.SerializeObject(Result);
            //Response.Write(JsonText);
        }
        else
        {
            var Result = new
            {
                SENDERACCOUNT = "ADMIN",
                SENDERPASSWORD = "ADMIN0424",
                SENDERTENANTID = "0081",
                HOSTINFO = "NewType",
                HOSTLANGUAGE = "TraditionalChinese",
                CASENO = CASENO,
                CASETYPE = CASETYPE,
                PAYMENTPERIOD = PAYMENTPERIOD,
                ISCONVERTPREPAYMENT = ISCONVERTPREPAYMENT,
                ALLOCATIONSTARTYEARMONTH = ALLOCATIONSTARTYEARMONTH,
                ALLOCATIONENDYEARMONTH = ALLOCATIONENDYEARMONTH,
		WITHHOLDINGINCOMETAXRATIO = WITHHOLDINGINCOMETAXRATIO,
                WITHHOLDINGINCOMEACCOUNTID = WITHHOLDINGINCOMEACCOUNTID,
                WITHHOLDINGINHEALTHINSURANCETAXRATIO = WITHHOLDINGINHEALTHINSURANCETAXRATIO,
                WITHHOLDINGINHEALTHINSURANCEPREMIUMACCOUNTID = WITHHOLDINGINHEALTHINSURANCEPREMIUMACCOUNTID,
                USECASE = USECASE,
                SERVICEID = "EXEPURCHASEPAYMENTENTRYBYAACQUIRE"
            };
            //序列化為JSON字串並輸出結果
            JsonText = JsonConvert.SerializeObject(Result);
            //Response.Write(JsonText);
        }
        
        
        //取FERP連線
        CommandText = "SELECT ConfigValue FROM SCSB_SystemConfig WHERE ConfigID=N'FERPServer' ";
        ParameterList.Clear();
        TmpDt = SqlQuery(connString, CommandText, ParameterList);
        if (TmpDt.Rows.Count > 0)
        {
            FErpServer = TmpDt.Rows[0]["ConfigValue"].ToString();
        }
        //呼叫erp
        //HttpWebRequest request = (HttpWebRequest)WebRequest.Create(FErpServer + @"/LCJsonBridge/JsonService.do");
        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(FErpServer + CallUrl);
        request.Method = "POST";
        request.ContentType = "application/json";
        StreamWriter requestWriter = new StreamWriter(request.GetRequestStream());

        try
        {
            requestWriter.Write(JsonText);
        }
        catch
        {
            throw;
        }
        finally
        {
            requestWriter.Close();
            requestWriter = null;
        }
        string responseStr = "";

        HttpWebResponse response = (HttpWebResponse)request.GetResponse();
        using (StreamReader sr = new StreamReader(response.GetResponseStream()))
        {
            responseStr = sr.ReadToEnd();
        }

        //Response.Write(responseStr);
        
        //解析分錄資料-
        try
        {
            JObject Jo = JObject.Parse(responseStr);
            dynamic dyna = Jo as dynamic;
            //檢查JSON格式
            if (dyna.RETURNVALUE.Value == "N")
            {
                ErrMsg += dyna.RETURNMSG.Value;
            }
        }
        catch (Exception ex)
        {
            Response.Write(OutputJosn(32, "ErpKey Json格式有誤", ""));
            Response.End();
        }
        //檢查Json內容是否正確
        if (ErrMsg != "")
        {
            Response.Write(OutputJosn(32, ErrMsg, ""));
            Response.End();
        }

        if (Type=="1")
        {
            JObject Jo2 = JObject.Parse(responseStr);

            //取分錄資料
            var ary = ((JArray)Jo2["SIMULATIONENTRYDATA"]).Cast<dynamic>().ToArray();
            DataText = "[" + string.Join(",", ary) + "]";

            //Response.Write(DataText);
            //解JsonToDT
            DetailDt = JsonConvert.DeserializeObject<DataTable>(DataText);
            if (DetailDt.Columns.Count > 0)
            {
                DetailDt.Columns.Add("RequisitionID");
                for (int i = 0; i < DetailDt.Rows.Count; i++)
                {
                    DetailDt.Rows[i]["RequisitionID"] = RequisitionID;
                }
            }
            //啟動交易機制-
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                SqlTransaction tran = conn.BeginTransaction();
                try
                {
                    TmpSql = "Delete from FM7T_A66_EXEPURCHASEPAYMENTENTRYBYAACQUIRE WHERE RequisitionID = @RequisitionID ";
                    using (SqlCommand cmd = new SqlCommand(TmpSql, conn, tran))
                    {
                        cmd.Parameters.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        int result = cmd.ExecuteNonQuery();
                    }

                    using (SqlBulkCopy sqlBC = new SqlBulkCopy(conn, SqlBulkCopyOptions.KeepNulls, tran))
                    {
                        sqlBC.BatchSize = 10;
                        //sqlBC.BulkCopyTimeout = 60;//超時之前操作完成所允許的秒數
                        sqlBC.DestinationTableName = "FM7T_A66_EXEPURCHASEPAYMENTENTRYBYAACQUIRE";
                        sqlBC.ColumnMappings.Add("RequisitionID", "RequisitionID");
                        sqlBC.ColumnMappings.Add("ENTITYID", "ENTITYID");
                        sqlBC.ColumnMappings.Add("DEBITORCREDIT", "DEBITORCREDIT");
                        sqlBC.ColumnMappings.Add("ACCOUNTID", "ACCOUNTID");
                        sqlBC.ColumnMappings.Add("ACCOUNTNAME", "ACCOUNTNAME");
                        sqlBC.ColumnMappings.Add("CURRENCYID", "CURRENCYID");
                        sqlBC.ColumnMappings.Add("CURRENCYNAME", "CURRENCYNAME");
                        sqlBC.ColumnMappings.Add("EXCHANGERATE", "EXCHANGERATE");
                        sqlBC.ColumnMappings.Add("TRANSACTIONAMOUNT", "TRANSACTIONAMOUNT");
                        sqlBC.ColumnMappings.Add("ENTRYAMOUNT", "ENTRYAMOUNT");
                        sqlBC.ColumnMappings.Add("DEPARTMENTID", "DEPARTMENTID");
                        sqlBC.ColumnMappings.Add("DEPARTMENTNAME", "DEPARTMENTNAME");
                        sqlBC.ColumnMappings.Add("VENDORUNIFORMINVOICENO", "VENDORUNIFORMINVOICENO");
                        sqlBC.ColumnMappings.Add("VENDORNAME", "VENDORNAME");
                        sqlBC.ColumnMappings.Add("SUMMARY", "SUMMARY");
                        sqlBC.ColumnMappings.Add("INVOICETYPE", "INVOICETYPE");
                        sqlBC.ColumnMappings.Add("INVOICENO", "INVOICENO");
                        sqlBC.ColumnMappings.Add("INVOICEDATE", "INVOICEDATE");
                        sqlBC.ColumnMappings.Add("INVOICECURRENCYID", "INVOICECURRENCYID");
                        sqlBC.ColumnMappings.Add("INVOICECURRENCYNAME", "INVOICECURRENCYNAME");
                        sqlBC.ColumnMappings.Add("INVOICEAMOUNT", "INVOICEAMOUNT");
                        sqlBC.ColumnMappings.Add("REMITTANCEBANKID", "REMITTANCEBANKID");
                        sqlBC.ColumnMappings.Add("REMITTANCEBANKNAME", "REMITTANCEBANKNAME");
                        sqlBC.ColumnMappings.Add("REMITTANCEBANKACCOUNT", "REMITTANCEBANKACCOUNT");
                        sqlBC.ColumnMappings.Add("REMITTANCEBANKACCOUNTNAME", "REMITTANCEBANKACCOUNTNAME");
                        sqlBC.WriteToServer(DetailDt);
                        tran.Commit();
                    }
                    Response.Write(OutputJosn(1, "驗收付款分錄取得成功", ""));
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Response.Write(OutputJosn(99, "驗收付款分錄寫入錯誤", ""));
                }
            }
        }
        else
        {
            Response.Write(OutputJosn(1, responseStr, ""));
        }
    }
    //輸出Json
    string OutputJosn(int Status, String ErrorMessage, string RequisitionID)
    {
        string JsonText = "";
        //用Linq直接組
        var Result = new
        {
            Status = Status,
            ErrorMessage = ErrorMessage,
            RequisitionID = RequisitionID
        };
        //序列化為JSON字串並輸出結果
        JsonText = JsonConvert.SerializeObject(Result);
        return JsonText;
    }
    //呼叫DB方式回傳DataTable
    public DataTable SqlQuery(string connString, string CommandString, List<SqlParameter> ParameterList)
    {
        try
        {
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(CommandString, conn))
                {
                    foreach (SqlParameter SP in ParameterList)
                    {
                        cmd.Parameters.Add(SP);
                    }

                    DataTable Reault = new DataTable();
                    Reault.Load(cmd.ExecuteReader());
                    Reault.Dispose();
                    conn.Close();

                    return Reault;
                }
            }
        }
        catch
        {
            return null;
        }
    }
    //呼叫DB方式不回傳
    public static bool SqlCommand(string connString, string CommandString, List<SqlParameter> ParameterList)
    {
        try
        {
            using (SqlConnection conn = new SqlConnection(connString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(CommandString, conn))
                {
                    foreach (SqlParameter SP in ParameterList)
                    {
                        cmd.Parameters.Add(SP);
                    }

                    cmd.ExecuteNonQuery();
                    conn.Close();
                }
            }
        }
        catch
        {
            return false;
        }
        return true;
    }

    //使用SqlConnection連線並傳回Datatable
    DataTable ExecSqlQuery(string connString, string CommandText)
    {
        DataTable dt = new DataTable();
        using (SqlConnection Conn = new SqlConnection(connString))
        {
            Conn.Open();
            SqlCommand cmd = new SqlCommand(CommandText, Conn);
            SqlDataAdapter dAdpter = new SqlDataAdapter(cmd);
            dAdpter.Fill(dt);
            Conn.Close();
        }
        return dt;
    }
    //使用SQL Connection連線僅傳回int
    int ExecSqlNonQuery(string connString, string CommandText)
    {
        int rVal;
        using (SqlConnection Conn = new SqlConnection(connString))
        {
            Conn.Open();
            SqlCommand cmd = new SqlCommand(CommandText, Conn);
            rVal = cmd.ExecuteNonQuery();
            Conn.Close();
        }
        return rVal;
    }
    //讀取連線字串
    string LoadCmdStr(String xdbcPath, int DBConType)
    {
        string FileName = "";
        string connectionString = "";
        const string userRoot = "HKEY_LOCAL_MACHINE";
        const string subkey = "Software\\NewType\\AutoWeb.Net";
        const string keyName = userRoot + "\\" + subkey;
        string Path = (string)Registry.GetValue(keyName, "Root", -1);
        Path = Path.Replace("\\", "\\\\");
        XdbcInfoIO objXdbc = new XdbcInfoIO();
        FileName = Path + xdbcPath;
        objXdbc.LoadFile(FileName, "");
        if (DBConType == 1)
        {
            connectionString = objXdbc.XdbcConnection.sMsSqlConnectString;
        }
        else
        {
            connectionString = objXdbc.XdbcConnection.sOleDBConnectString;
        }
        return connectionString;
    }
}