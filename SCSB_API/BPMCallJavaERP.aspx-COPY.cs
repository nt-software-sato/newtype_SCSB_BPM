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

public partial class BPMCallJavaERP : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        string RequisitionID = HttpUtility.HtmlEncode(Request.QueryString["RequisitionID"]) ?? string.Empty;
        string Status = HttpUtility.HtmlEncode(Request.QueryString["Status"]) ?? string.Empty;

        //宣告要ERP起單的表單名稱
        string JsonText = "";
        string CommandText = "";
        string FErpServer = "";

        DataTable TmpDt = new DataTable();
        DataTable dt = new DataTable();
        List<SqlParameter> ParameterList = new List<SqlParameter>();

        //RequisitionID = "2a9a48f0-aa1c-4d87-87bd-0eee6b730dfc";
        //使用AutoWeb資料庫連線字串
        string connString = LoadCmdStr("\\\\Database\\\\Project\\\\SCSB\\\\Flow\\\\connection\\\\BPM.xdbc.xmf", 1);

        //使用Sql連線並傳回Datatable
        CommandText =
            " Select R.Status, Diagram.Identify " +
            "   from FSe7en_Sys_Requisition R inner join FSe7en_Sys_DiagramList Diagram " +
            "     on R.DiagramID=Diagram.DiagramID "+
            "  where RequisitionID=@RequisitionID ";

        ParameterList.Clear();
        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));

        dt = SqlQuery(connString, CommandText, ParameterList);

        if (dt.Rows.Count > 0)
        {
            //取主表資料-
            CommandText =
            " Select ENTITYID, DATAVERSION " +
            "   from FM7T_" + dt.Rows[0]["Identify"].ToString() + "_M " +
            "  where RequisitionID=@RequisitionID ";

            ParameterList.Clear();
            ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));

            TmpDt = SqlQuery(connString, CommandText, ParameterList);
            //用Linq直接組Json
            var Result = new
            {
                HOSTLANGUAGE = "TraditionalChinese",
                HOSTINFO = "NewType",
                SENDERACCOUNT = "ADMIN",
                SENDERPASSWORD = "ADMIN0424",
                SENDERTENANTID = "0081",
                USECASE = "UC_EM_ENTRYWHOLEBATCH",
                SERVICEID = "EXEASSETADJUSTENTRYANDBPMCREATE",
                ENTITYID = TmpDt.Rows[0]["ENTITYID"].ToString(),
                DATAVERSION = int.Parse(TmpDt.Rows[0]["DATAVERSION"].ToString()),
                NEWTYPEBPMREQUISITIONID = "",
                NEWTYPEBPMSTATUSTYPE = Status
            };
            //序列化為JSON字串並輸出結果
            JsonText = JsonConvert.SerializeObject(Result);

            //Response.Write(JsonText);
            //取FERP連線
            CommandText = "SELECT ConfigValue FROM SCSB_SystemConfig WHERE ConfigID=N'FERPServer' ";
            ParameterList.Clear();
            TmpDt = SqlQuery(connString, CommandText, ParameterList);
            if (TmpDt.Rows.Count > 0)
            {
                FErpServer = TmpDt.Rows[0]["ConfigValue"].ToString();
            }
            //呼叫erp
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(FErpServer + @"/LCJsonBridge/JsonService.do");
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

            Response.Write(responseStr);
        }
        
        
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