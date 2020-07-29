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
using System.Text.RegularExpressions;

public partial class FERPCallBPM : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //宣告要ERP起單的表單名稱
        //Request參數 Request 接在網址列?後方的參數
        string ErpKey = HttpUtility.HtmlEncode(Request.Form["ErpKey"]).Replace("&quot;", "\"").Replace("&#39;", "'") ?? string.Empty;
        //string ErpKey = "";
        string Identify = "";
        string AccountID = "";
        string DeptID = "";
        string CounterSignID = "";
        string CounterSignName = "";
        string OrgApplicantID = "";
        string JsonText = "", RowData = "";
        string TmpSql = "";
        string ErrMsg = "";
        //用datatable來裝JSON值，免去擴充時要不斷宣告類別檔
        DataTable TmpDt = new DataTable();
        DataTable Dt = new DataTable(); 
        List<SqlParameter> ParameterList = new List<SqlParameter>();

        //使用AutoWeb資料庫連線字串
        string connString = LoadCmdStr("\\\\Database\\\\Project\\\\SCSB\\\\Flow\\\\connection\\\\BPM.xdbc.xmf", 1);
        //透過LoadCmdStr 將字串轉為資料庫連字串
        JsonText = HttpUtility.HtmlEncode(ErpKey).Replace("&quot;", "\"").Replace("&#39;", "'"); //主要資安要轉字串

        ////取Json字串
        //FileStream fs = new FileStream(Server.MapPath("~/FJsonText.txt"), FileMode.Open);
        //StreamReader sr = new StreamReader(fs);
        ////載入JSON字串
        //JsonText = sr.ReadToEnd();
        //sr.Close();
        //fs.Close();

        //解JsonToDT
        if (JsonText == "")
        {
            Response.Write(OutputJosn(32, "ErpKey 未輸入", ""));
        }
        else
        {
            try
            {
                JsonText = ReplaceStr(JsonText);
                //JsonText = "[" + JsonText + "]";
                JObject Jo = JObject.Parse(JsonText);
                dynamic dyna = Jo as dynamic;

                //檢查JSON格式
                if (dyna.Identify == null || dyna.AccountID == null || dyna.DeptID == null || dyna.RowData == null)
                {
                    //Response.Write(OutputJosn(32, "ErpKey Json格式有誤", ""));
                    ErrMsg += "ErpKey Json格式有誤";
                }

                Identify = dyna.Identify;
                AccountID = dyna.AccountID;
                DeptID = dyna.DeptID;

                if (dyna.CounterSignID != null)
                {
                    CounterSignID = dyna.CounterSignID;
                    CounterSignName = dyna.CounterSignName;
                }
                var ary = ((JArray)Jo["RowData"]).Cast<dynamic>().ToArray();
                RowData = string.Join(",", ary);
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
            //檢查申請人是否存在
            TmpSql = "SELECT Enabled FROM FSe7en_Org_MemberStruct WHERE AccountID=@AccountID AND DeptID=@DeptID ";
            ParameterList.Clear();
            ParameterList.Add(new SqlParameter("@AccountID", AccountID.Replace("'", "")));
            ParameterList.Add(new SqlParameter("@DeptID", DeptID.Replace("'", "")));
            TmpDt = SqlQuery(connString, TmpSql, ParameterList);
            if (TmpDt.Rows.Count == 0)
            {
                Response.Write(OutputJosn(32, "申請人帳號不存在", ""));
                Response.End();
            }
            else
            {
                OrgApplicantID = AccountID;
                //為未啟用帳號取該部門虛擬帳號
                if (TmpDt.Rows[0]["Enabled"].ToString() == "0")
                {
                    TmpSql = "SELECT AccountID FROM FSe7en_Org_MemberStruct WHERE AccountID=@AccountID and DeptID=@DeptID ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@AccountID", "D_" + DeptID.Replace("'", "")));
                    ParameterList.Add(new SqlParameter("@DeptID", DeptID.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        JsonText = OutputJosn(32, "部門虛擬帳號不存在", "");
                        Response.Write(JsonText);
                        Response.End();
                    }
                    else
                    {
                        AccountID = TmpDt.Rows[0]["AccountID"].ToString();
                    }
                }
                if (Identify != "")
                {
                    //取得RequisitionID
                    string RequisitionID = Guid.NewGuid().ToString();
                    Response.Write(StarFlow(connString, RequisitionID, RowData, Identify.Replace("'", ""), AccountID.Replace("'", ""), DeptID.Replace("'", ""), OrgApplicantID.Replace("'", ""), CounterSignID.Replace("'", ""), CounterSignName.Replace("'", "")));
                }
            }
        }
    }

    //讀取連線字串
    string LoadCmdStr(String xdbcPath, int DBConType)
    {
        string FileName = "";
        string connectionString = "";
        const string userRoot = "HKEY_LOCAL_MACHINE"; //定義常數
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
    //BPM起單
    string StarFlow(string connString, string RequisitionID, string RowData, string Identify, string ApplicantID, string ApplicantDept, string OrgApplicantID, string CounterSignID, string CounterSignName)
    {
        string ErrorMsg = "";
        DataTable TmpDt = new DataTable();
        DataTable DetailDt = new DataTable();
        string ITable = Identify;
        string FlowID = Identify;
        string ApplicantName = "";
        string ApplicantDeptName = "";
        string OrgApplicantName = "";
        string DiagramID = "";
        string TmpSql = "";
        string DBName = "";
        string DataText = "";
        int CounterSignCount = 0;
        /*20200130修正
        string OrgRequisitionID = "";
        */
        List<SqlParameter> ParameterList = new List<SqlParameter>();
        List<String> StrList1 = new List<String>();
        List<String> StrList2 = new List<String>();

        //計算會簽單位數量
        if (CounterSignID != "")
        {
            string[] CounterSignArray = CounterSignID.Split(',');
            CounterSignCount = CounterSignArray.Length;
        }

        Connector FM7Engine = new Connector();
        if (!FM7Engine.Initialize("SCSB_Flow", ref ErrorMsg)) //呼叫實質型別的無參數建構函式，初始化陣列的每個項目
        {
            return ErrorMsg;
        }

        //取ERP連線
        TmpSql = "SELECT ConfigValue FROM SCSB_SystemConfig WHERE ConfigID=N'ERPServer' ";
        ParameterList.Clear(); 
        TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        if (TmpDt.Rows.Count > 0)
        {
            DBName = TmpDt.Rows[0]["ConfigValue"].ToString();
        }

        //取申請人名稱及部門名稱
        TmpSql = "SELECT MemberName, DeptName FROM F7Organ_View_CurrMember WHERE AccountID=@AccountID and DeptID=@DeptID ";
        ParameterList.Clear();
        ParameterList.Add(new SqlParameter("@AccountID", ApplicantID.Replace("'", "")));
        ParameterList.Add(new SqlParameter("@DeptID", ApplicantDept.Replace("'", "")));
        TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        if (TmpDt.Rows.Count > 0)
        {
            ApplicantName = TmpDt.Rows[0]["MemberName"].ToString();
            ApplicantDeptName = TmpDt.Rows[0]["DeptName"].ToString();
        }
        //取原申請人名稱
        TmpSql = "SELECT DisplayName FROM FSe7en_Org_MemberInfo WHERE AccountID=@AccountID ";
        ParameterList.Clear();
        ParameterList.Add(new SqlParameter("@AccountID", OrgApplicantID.Replace("'", "")));
        TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        if (TmpDt.Rows.Count > 0)
        {
            OrgApplicantName = TmpDt.Rows[0]["DisplayName"].ToString();
        }
        //取流程代號
        TmpSql = "SELECT DiagramID FROM FSe7en_Sys_DiagramList WHERE Identify=@Identify ";
        ParameterList.Clear();
        ParameterList.Add(new SqlParameter("@Identify", Identify.Replace("'", "")));
        TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        if (TmpDt.Rows.Count > 0)
        {
            DiagramID = TmpDt.Rows[0]["DiagramID"].ToString();
        }
        DiagramGuid sDiagramGuid = new DiagramGuid(DiagramID);

        //取M表資料
        DataTable MTable = new DataTable();
        MTable.Columns.Add(new DataColumn("RequisitionID"));
        MTable.Columns.Add(new DataColumn("DiagramID"));
        MTable.Columns.Add(new DataColumn("ApplicantDept"));
        MTable.Columns.Add(new DataColumn("ApplicantDeptName"));
        MTable.Columns.Add(new DataColumn("ApplicantID"));
        MTable.Columns.Add(new DataColumn("ApplicantName"));
        MTable.Columns.Add(new DataColumn("FillerID"));
        MTable.Columns.Add(new DataColumn("FillerName"));
        MTable.Columns.Add(new DataColumn("ApplicantDateTime"));
        MTable.Columns.Add(new DataColumn("Priority"));
        MTable.Columns.Add(new DataColumn("DraftFlag"));
        if (Identify != "F35" && Identify != "F36")
        {
            MTable.Columns.Add(new DataColumn("CounterSignID"));
            MTable.Columns.Add(new DataColumn("CounterSignName"));
            MTable.Columns.Add(new DataColumn("SubProcessCount"));
            MTable.Columns.Add(new DataColumn("ExecSubProcess"));
        }
        MTable.Columns.Add(new DataColumn("OrgApplicantID"));
        MTable.Columns.Add(new DataColumn("OrgApplicantName"));
        //20190524修改
        //MTable.Columns.Add(new DataColumn("ENTITYID"));
        //MTable.Columns.Add(new DataColumn("DATAVERSION"));
        //MTable.Columns.Add(new DataColumn("ASSETAPPLICATIONADJUSTID"));
        MTable.Columns.Add(new DataColumn("CURRENTUSEDEPARTMENTID"));
        MTable.Columns.Add(new DataColumn("CURRENTUSEDEPARTMENTNAME"));
        MTable.Columns.Add(new DataColumn("CHAGEUSEDEPARTMENTID"));
        MTable.Columns.Add(new DataColumn("CHAGEUSEDEPARTMENTNAME"));
        MTable.Columns.Add(new DataColumn("REMARK"));
        MTable.Columns.Add(new DataColumn("ErpFlag"));
        DataRow M_row = MTable.NewRow();

        JObject Jo = JObject.Parse(RowData);
        dynamic dyna = Jo as dynamic;
        //拆解Detail
        var ary = ((JArray)Jo["SIMULATIONAPPLICATIONPASSASSETCONDITIONDATA"]).Cast<dynamic>().ToArray();

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
        M_row["RequisitionID"] = RequisitionID;
        M_row["DiagramID"] = DiagramID;
        M_row["ApplicantDept"] = ApplicantDept;
        M_row["ApplicantDeptName"] = ApplicantDeptName;
        M_row["ApplicantID"] = ApplicantID;
        M_row["ApplicantName"] = ApplicantName;
        M_row["FillerID"] = ApplicantID;
        M_row["FillerName"] = ApplicantName;
        M_row["ApplicantDateTime"] = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
        M_row["Priority"] = 2;
        M_row["DraftFlag"] = 0;
        M_row["ErpFlag"] = "0";

        if (Identify != "F35" && Identify != "F36")
        {
            M_row["CounterSignID"] = CounterSignID;
            M_row["CounterSignName"] = CounterSignName;
            M_row["SubProcessCount"] = CounterSignCount;
            M_row["ExecSubProcess"] = 0;
        }
        /*20200130修正
        if (dyna.OrgRequisitionID != null)
        {
            OrgRequisitionID = dyna.OrgRequisitionID;
        }
        */
        M_row["OrgApplicantID"] = OrgApplicantID;
        M_row["OrgApplicantName"] = OrgApplicantName;
        /*
        M_row["ENTITYID"] = dyna.ENTITYID;
        M_row["DATAVERSION"] = dyna.DATAVERSION;
         */
        //M_row["ASSETAPPLICATIONADJUSTID"] = dyna.ASSETAPPLICATIONADJUSTID;
        M_row["REMARK"] = dyna.REMARK;
        //取SIMULATIONAPPLICATIONPASSASSETCONDITIONDATA內的值給M表用
        M_row["CURRENTUSEDEPARTMENTID"] = DetailDt.Rows[0]["CURRENTUSEDEPARTMENTID"].ToString();
        M_row["CURRENTUSEDEPARTMENTNAME"] = DetailDt.Rows[0]["CURRENTUSEDEPARTMENTNAME"].ToString();
        M_row["CHAGEUSEDEPARTMENTID"] = DetailDt.Rows[0]["CHAGEUSEDEPARTMENTID"].ToString();
        M_row["CHAGEUSEDEPARTMENTNAME"] = DetailDt.Rows[0]["CHAGEUSEDEPARTMENTNAME"].ToString();
        MTable.Rows.Add(M_row);

        //將資料寫入該表單M與D表
        switch (Identify)
        {
            case "F31":
            case "F32":
            case "F33":
            case "F34":
            case "F35":
            case "F36":
                //判斷是否有值
                if (DetailDt.Rows.Count == 0)
                {
                    ErrorMsg = OutputJosn(99, "來源明細沒資料", "");
                    return ErrorMsg;
                }
                else
                {
                    using (SqlConnection conn = new SqlConnection(connString))
                    {
                        conn.Open();
                        SqlTransaction tran = conn.BeginTransaction();
                        try
                        {
                            using (SqlBulkCopy sqlBC = new SqlBulkCopy(conn, SqlBulkCopyOptions.KeepNulls, tran))
                            {
                                sqlBC.BatchSize = 10;
                                //sqlBC.BulkCopyTimeout = 60;//超時之前操作完成所允許的秒數
                                sqlBC.DestinationTableName = "FM7T_" + Identify + "_M";
                                sqlBC.ColumnMappings.Add("RequisitionID", "RequisitionID");
                                sqlBC.ColumnMappings.Add("DiagramID", "DiagramID");
                                sqlBC.ColumnMappings.Add("ApplicantDept", "ApplicantDept");
                                sqlBC.ColumnMappings.Add("ApplicantDeptName", "ApplicantDeptName");
                                sqlBC.ColumnMappings.Add("ApplicantID", "ApplicantID");
                                sqlBC.ColumnMappings.Add("ApplicantName", "ApplicantName");
                                sqlBC.ColumnMappings.Add("FillerID", "FillerID");
                                sqlBC.ColumnMappings.Add("FillerName", "FillerName");
                                sqlBC.ColumnMappings.Add("ApplicantDateTime", "ApplicantDateTime");
                                sqlBC.ColumnMappings.Add("Priority", "Priority");
                                sqlBC.ColumnMappings.Add("DraftFlag", "DraftFlag");
								
                                if (Identify != "F35" && Identify != "F36")
                                {
                                    sqlBC.ColumnMappings.Add("CounterSignID", "CounterSignID");
                                    sqlBC.ColumnMappings.Add("CounterSignName", "CounterSignName");
                                    sqlBC.ColumnMappings.Add("SubProcessCount", "SubProcessCount");
                                    sqlBC.ColumnMappings.Add("ExecSubProcess", "ExecSubProcess");
                                }
                                sqlBC.ColumnMappings.Add("OrgApplicantID", "OrgApplicantID");
                                sqlBC.ColumnMappings.Add("OrgApplicantName", "OrgApplicantName");
                                /*
                                sqlBC.ColumnMappings.Add("ENTITYID", "ENTITYID");
                                sqlBC.ColumnMappings.Add("DATAVERSION", "DATAVERSION");
                                */
                                //sqlBC.ColumnMappings.Add("ASSETAPPLICATIONADJUSTID", "ASSETAPPLICATIONADJUSTID");
                                sqlBC.ColumnMappings.Add("REMARK", "REMARK");
                                sqlBC.ColumnMappings.Add("CURRENTUSEDEPARTMENTID", "CURRENTUSEDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("CURRENTUSEDEPARTMENTNAME", "CURRENTUSEDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("CHAGEUSEDEPARTMENTID", "CHAGEUSEDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("CHAGEUSEDEPARTMENTNAME", "CHAGEUSEDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("ErpFlag", "ErpFlag");
                                sqlBC.WriteToServer(MTable);								
								
                            }
                            using (SqlBulkCopy sqlBC = new SqlBulkCopy(conn, SqlBulkCopyOptions.KeepNulls, tran))
                            {
                                sqlBC.BatchSize = 200;
                                //sqlBC.BulkCopyTimeout = 60;//超時之前操作完成所允許的秒數                               
                                sqlBC.DestinationTableName = "FM7T_" + Identify + "_EACHSITEASSETPERIODCHANGEATA";                               
                                sqlBC.ColumnMappings.Add("RequisitionID", "RequisitionID");
                                //20190729_DB欄位重新命名
                                //sqlBC.ColumnMappings.Add("CURRENTASSETNAME", "CURRENTASSETNAME");
                                //sqlBC.ColumnMappings.Add("CHANGECURRENTJOURNALAMOUNT", "CHANGECURRENTJOURNALAMOUNT");
                                //sqlBC.ColumnMappings.Add("CHANGEACQUIREDATE", "CHANGEACQUIREDATE");
                                //sqlBC.ColumnMappings.Add("CHANGEDURATIONYEARS", "CHANGEDURATIONYEARS");
                                //sqlBC.ColumnMappings.Add("CHANGEACCUMULATEDEPRECIATEAMOUNT", "CHANGEACCUMULATEDEPRECIATEAMOUNT");
                                //sqlBC.ColumnMappings.Add("CHANGEBRAND", "CHANGEBRAND");
                                //sqlBC.ColumnMappings.Add("CHANGEMODEL", "CHANGEMODEL");
                                //sqlBC.ColumnMappings.Add("CHANGESPECIFICATION", "CHANGESPECIFICATION");
                                //sqlBC.ColumnMappings.Add("CHANGESALVAGEVALUE", "CHANGESALVAGEVALUE");
                                //sqlBC.ColumnMappings.Add("SALESAMOUNT", "SALESAMOUNT");
                                //sqlBC.ColumnMappings.Add("INCOMEAMOUNT", "INCOMEAMOUNT");
                                //sqlBC.ColumnMappings.Add("CURRENTDURATIONYEARS", "CURRENTDURATIONYEARS");
                                //sqlBC.ColumnMappings.Add("CURRENTINCREASEREDUCEDURATIONMONTHS", "CURRENTINCREASEREDUCEDURATIONMONTHS");
                                //sqlBC.ColumnMappings.Add("CURRENTACCUMULATEDEPRECIATEMONTHS", "CURRENTACCUMULATEDEPRECIATEMONTHS");
                                //sqlBC.ColumnMappings.Add("CHANGEACQUIRECOST", "CHANGEACQUIRECOST");
                                //sqlBC.ColumnMappings.Add("CHANGEINCREASEREDUCEDURATIONMONTHS", "CHANGEINCREASEREDUCEDURATIONMONTHS");
                                //sqlBC.ColumnMappings.Add("CHANGEACCUMULATEDEPRECIATEMONTHS", "CHANGEACCUMULATEDEPRECIATEMONTHS");
                                sqlBC.ColumnMappings.Add("ASSETID", "ASSETID");
                                sqlBC.ColumnMappings.Add("ASSETNAME", "ASSETNAME");
                                sqlBC.ColumnMappings.Add("CURRENTINCREASEREDUCEACQUIRECOST", "CURRENTINCREASEREDUCEACQUIRECOST");
                                sqlBC.ColumnMappings.Add("CURRENTSALVAGEVALUE", "CURRENTSALVAGEVALUE");
                                sqlBC.ColumnMappings.Add("CURRENTDURATIONYEARS", "CURRENTDURATIONYEARS");
                                sqlBC.ColumnMappings.Add("CURRENTINCREASEREDUCEDURATIONMONTHS", "CURRENTINCREASEREDUCEDURATIONMONTHS");
                                sqlBC.ColumnMappings.Add("CURRENTACCUMULATEDEPRECIATEMONTHS", "CURRENTACCUMULATEDEPRECIATEMONTHS");
                                sqlBC.ColumnMappings.Add("CURRENTUSEDEPARTMENTID", "CURRENTUSEDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("CURRENTUSEDEPARTMENTNAME", "CURRENTUSEDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("CHAGEACQUIREDATE", "CHAGEACQUIREDATE");
                                sqlBC.ColumnMappings.Add("CHAGEACQUIRECOST", "CHAGEACQUIRECOST");
                                sqlBC.ColumnMappings.Add("CHAGESALVAGEVALUE", "CHAGESALVAGEVALUE");
                                sqlBC.ColumnMappings.Add("CHAGEDURATIONYEARS", "CHAGEDURATIONYEARS");
                                sqlBC.ColumnMappings.Add("CHAGEINCREASEREDUCEDURATIONMONTHS", "CHAGEINCREASEREDUCEDURATIONMONTHS");
                                sqlBC.ColumnMappings.Add("CHAGEACCUMULATEDEPRECIATEMONTHS", "CHAGEACCUMULATEDEPRECIATEMONTHS");
                                sqlBC.ColumnMappings.Add("CHAGEACCUMULATEDEPRECIATEAMOUNT", "CHAGEACCUMULATEDEPRECIATEAMOUNT");
                                sqlBC.ColumnMappings.Add("CHAGEBRAND", "CHAGEBRAND");
                                sqlBC.ColumnMappings.Add("CHAGEMODEL", "CHAGEMODEL");
                                sqlBC.ColumnMappings.Add("CHAGESPECIFICATION", "CHAGESPECIFICATION");
                                sqlBC.ColumnMappings.Add("CHAGEUSEDEPARTMENTID", "CHAGEUSEDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("CHAGEUSEDEPARTMENTNAME", "CHAGEUSEDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("CHAGECURRENTJOURNALAMOUNT", "CHAGECURRENTJOURNALAMOUNT");
                                sqlBC.ColumnMappings.Add("ASSETSALESAMOUNT", "ASSETSALESAMOUNT");
                                sqlBC.ColumnMappings.Add("ASSETINCOMEAMOUNT", "ASSETINCOMEAMOUNT");

                                /*sqlBC.ColumnMappings.Add("CURRENTASSETNAME", "CURRENTASSETNAME"); 20200130修正
                                sqlBC.ColumnMappings.Add("CURRENTASSETUMNAME", "CURRENTASSETUMNAME");
                                sqlBC.ColumnMappings.Add("ASSETLOCATIONID", "ASSETLOCATIONID");
                                sqlBC.ColumnMappings.Add("ASSETLOCATIONNAME", "ASSETLOCATIONNAME");
                                sqlBC.ColumnMappings.Add("ASSETGROUP", "ASSETGROUP");
                                sqlBC.ColumnMappings.Add("ASSETUMID", "ASSETUMID");
                                sqlBC.ColumnMappings.Add("ASSETUMNAME", "ASSETUMNAME");
                                sqlBC.ColumnMappings.Add("ASSETTYPEID", "ASSETTYPEID");
                                sqlBC.ColumnMappings.Add("ASSETTYPENAME", "ASSETTYPENAME");
                                sqlBC.ColumnMappings.Add("ASSETSALESTYPE", "ASSETSALESTYPE");
                                sqlBC.ColumnMappings.Add("ASSETDISPOSALADJUSTELEMENT", "ASSETDISPOSALADJUSTELEMENT");
                                sqlBC.ColumnMappings.Add("ASSETAMOUNTADJUSTADJUSTELEMENT", "ASSETAMOUNTADJUSTADJUSTELEMENT");
                                sqlBC.ColumnMappings.Add("ASSETQUANTITY", "ASSETQUANTITY");
                                sqlBC.ColumnMappings.Add("ADJUSTTYPE", "ADJUSTTYPE");
                                sqlBC.ColumnMappings.Add("BRAND", "BRAND");
                                sqlBC.ColumnMappings.Add("CHAGEASSETLOCATIONID", "CHAGEASSETLOCATIONID");
                                sqlBC.ColumnMappings.Add("CHAGEASSETLOCATIONNAME", "CHAGEASSETLOCATIONNAME");
                                sqlBC.ColumnMappings.Add("CHAGEMANAGEMENTDEPARTMENTID", "CHAGEMANAGEMENTDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("CHAGEMANAGEMENTDEPARTMENTNAME", "CHAGEMANAGEMENTDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("CHAGEUSEEMPLOYEEID", "CHAGEUSEEMPLOYEEID");
                                sqlBC.ColumnMappings.Add("CHAGEUSEEMPLOYEENAME", "CHAGEUSEEMPLOYEENAME");   
                                sqlBC.ColumnMappings.Add("CHAGEASSETUMID", "CHAGEASSETUMID");
                                sqlBC.ColumnMappings.Add("CHAGEASSETUMNAME", "CHAGEASSETUMNAME");
                                sqlBC.ColumnMappings.Add("CHAGEMANAGEMENTEMPLOYEEID", "CHAGEMANAGEMENTEMPLOYEEID");
                                sqlBC.ColumnMappings.Add("CHAGEMANAGEMENTEMPLOYEENAME", "CHAGEMANAGEMENTEMPLOYEENAME");
                                sqlBC.ColumnMappings.Add("CHAGECONSUMEUMID", "CHAGECONSUMEUMID");
                                sqlBC.ColumnMappings.Add("CHAGECONSUMEUMNAME", "CHAGECONSUMEUMNAME");
                                sqlBC.ColumnMappings.Add("CHAGEASSETQUANTITY", "CHAGEASSETQUANTITY");
                                sqlBC.ColumnMappings.Add("CHAGEONASSETBOOKORNOTONASSETBOOK", "CHAGEONASSETBOOKORNOTONASSETBOOK");
                                sqlBC.ColumnMappings.Add("CHAGEDEPRECIATECOMPLETEDYEARMONTH", "CHAGEDEPRECIATECOMPLETEDYEARMONTH"); 
                                sqlBC.ColumnMappings.Add("CHAGEOWNERSHIPSTATUS", "CHAGEOWNERSHIPSTATUS");
                                sqlBC.ColumnMappings.Add("CHAGEDEPRECIATEMETHOD", "CHAGEDEPRECIATEMETHOD");
                                sqlBC.ColumnMappings.Add("CHAGESTARTDEPRECIATEYEARMONTH", "CHAGESTARTDEPRECIATEYEARMONTH");
                                sqlBC.ColumnMappings.Add("CHAGEDEPRECIATESTATUS", "CHAGEDEPRECIATESTATUS");
                                sqlBC.ColumnMappings.Add("CHAGEINCREASEREDUCEACQUIRECOST", "CHAGEINCREASEREDUCEACQUIRECOST");
                                sqlBC.ColumnMappings.Add("CONSUMEUMID", "CONSUMEUMID");
                                sqlBC.ColumnMappings.Add("CONSUMEUMNAME", "CONSUMEUMNAME");
                                sqlBC.ColumnMappings.Add("CURRENTMANAGEMENTDEPARTMENTID", "CURRENTMANAGEMENTDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("CURRENTMANAGEMENTDEPARTMENTNAME", "CURRENTMANAGEMENTDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("CURRENTMANAGEMENTEMPLOYEEID", "CURRENTMANAGEMENTEMPLOYEEID");
                                sqlBC.ColumnMappings.Add("CURRENTMANAGEMENTEMPLOYEENAME", "CURRENTMANAGEMENTEMPLOYEENAME");
                                sqlBC.ColumnMappings.Add("CURRENTUSEEMPLOYEEID", "CURRENTUSEEMPLOYEEID");
                                sqlBC.ColumnMappings.Add("CURRENTUSEEMPLOYEENAME", "CURRENTUSEEMPLOYEENAME");
                                sqlBC.ColumnMappings.Add("CURRENTCONSUMEUMID", "CURRENTCONSUMEUMID");
                                sqlBC.ColumnMappings.Add("CURRENTCONSUMEUMNAME", "CURRENTCONSUMEUMNAME");
                                sqlBC.ColumnMappings.Add("CURRENTASSETLOCATIONID", "CURRENTASSETLOCATIONID");
                                sqlBC.ColumnMappings.Add("CURRENTASSETLOCATIONNAME", "CURRENTASSETLOCATIONNAME");
                                sqlBC.ColumnMappings.Add("CURRENTASSETUMID", "CURRENTASSETUMID");
                                sqlBC.ColumnMappings.Add("CURRENTASSETQUANTITY", "CURRENTASSETQUANTITY");
                                sqlBC.ColumnMappings.Add("CURRENTDEPRECIATESTATUS", "CURRENTDEPRECIATESTATUS");
                                sqlBC.ColumnMappings.Add("CURRENTOWNERSHIPSTATUS", "CURRENTOWNERSHIPSTATUS");
                                sqlBC.ColumnMappings.Add("CURRENTCURRENTJOURNALAMOUNT", "CURRENTCURRENTJOURNALAMOUNT");
                                sqlBC.ColumnMappings.Add("CURRENTBRAND", "CURRENTBRAND");
                                sqlBC.ColumnMappings.Add("CURRENTMODEL", "CURRENTMODEL");
                                sqlBC.ColumnMappings.Add("CURRENTONASSETBOOKORNOTONASSETBOOK", "CURRENTONASSETBOOKORNOTONASSETBOOK");
                                sqlBC.ColumnMappings.Add("CURRENTACQUIRECOST", "CURRENTACQUIRECOST");
                                sqlBC.ColumnMappings.Add("CURRENTACCUMULATEDEPRECIATEAMOUNT", "CURRENTACCUMULATEDEPRECIATEAMOUNT");
                                sqlBC.ColumnMappings.Add("CURRENTSPECIFICATION", "CURRENTSPECIFICATION");
                                sqlBC.ColumnMappings.Add("CURRENTSTARTDEPRECIATEYEARMONTH", "CURRENTSTARTDEPRECIATEYEARMONTH");
                                sqlBC.ColumnMappings.Add("CURRENTDEPRECIATEMETHOD", "CURRENTDEPRECIATEMETHOD");
                                sqlBC.ColumnMappings.Add("CURRENTACQUIREDATE", "CURRENTACQUIREDATE");
                                sqlBC.ColumnMappings.Add("CURRENTDEPRECIATECOMPLETEDYEARMONTH", "CURRENTDEPRECIATECOMPLETEDYEARMONTH");
                                sqlBC.ColumnMappings.Add("ENTITYID", "ENTITYID");
                                sqlBC.ColumnMappings.Add("ENTITYNAME", "ENTITYNAME");
                                sqlBC.ColumnMappings.Add("INCREASEREDUCEACQUIRECOST", "INCREASEREDUCEACQUIRECOST");
                                sqlBC.ColumnMappings.Add("INCREASEREDUCEDURATIONMONTHS", "INCREASEREDUCEDURATIONMONTHS");
                                sqlBC.ColumnMappings.Add("ISDISPOSALMONTHCALDEPRECIATE", "ISDISPOSALMONTHCALDEPRECIATE");
                                sqlBC.ColumnMappings.Add("JOURNALIZESITEID", "JOURNALIZESITEID");
                                sqlBC.ColumnMappings.Add("JOURNALIZESITENAME", "JOURNALIZESITENAME");
                                sqlBC.ColumnMappings.Add("MODEL", "MODEL");
                                sqlBC.ColumnMappings.Add("MANAGEMENTDEPARTMENTID", "MANAGEMENTDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("MANAGEMENTDEPARTMENTNAME", "MANAGEMENTDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("MANAGEMENTEMPLOYEEID", "MANAGEMENTEMPLOYEEID");
                                sqlBC.ColumnMappings.Add("MANAGEMENTEMPLOYEENAME", "MANAGEMENTEMPLOYEENAME");
                                sqlBC.ColumnMappings.Add("OPERATIONDATE", "OPERATIONDATE");
                                sqlBC.ColumnMappings.Add("OPERATOR", "OPERATOR");
                                sqlBC.ColumnMappings.Add("OWNERSHIPSTATUS", "OWNERSHIPSTATUS");
                                sqlBC.ColumnMappings.Add("SPECIFICATION", "SPECIFICATION");
                                sqlBC.ColumnMappings.Add("SALVAGEVALUE", "SALVAGEVALUE");
                                sqlBC.ColumnMappings.Add("SEQUENCENO", "SEQUENCENO");
                                sqlBC.ColumnMappings.Add("SYSID", "SYSID");
                                sqlBC.ColumnMappings.Add("TSTAMP", "TSTAMP");
                                sqlBC.ColumnMappings.Add("TRANSFERINUSEDEPARTMENTID", "TRANSFERINUSEDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("TRANSFERINUSEDEPARTMENTNAME", "TRANSFERINUSEDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("TRANSFERINENTITYID", "TRANSFERINENTITYID");
                                sqlBC.ColumnMappings.Add("TRANSFERINENTITYNAME", "TRANSFERINENTITYNAME");
                                sqlBC.ColumnMappings.Add("TRANSFERINMANAGEMENTDEPARTMENTID", "TRANSFERINMANAGEMENTDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("TRANSFERINMANAGEMENTDEPARTMENTNAME", "TRANSFERINMANAGEMENTDEPARTMENTNAME");
                                sqlBC.ColumnMappings.Add("USEEMPLOYEEID", "USEEMPLOYEEID");
                                sqlBC.ColumnMappings.Add("USEEMPLOYEENAME", "USEEMPLOYEENAME");
                                sqlBC.ColumnMappings.Add("USEDEPARTMENTID", "USEDEPARTMENTID");
                                sqlBC.ColumnMappings.Add("USEDEPARTMENTNAME", "USEDEPARTMENTNAME");   
                                */
                                sqlBC.WriteToServer(DetailDt); //DetailDt D表
                                tran.Commit();
	
                            }
                        }
                        catch (Exception ex)
                        {  
                            tran.Rollback();
                            return OutputJosn(99, "寫入發生錯誤請檢查Json資料是否正確", "");
                        }
                    }

                }

                break;
        }

        //起單跑流程
        using (SqlConnection InterfaceConnection = new SqlConnection(connString))
        {
            InterfaceConnection.Open();
            if (FM7Engine.Start(InterfaceConnection, RequisitionID, sDiagramGuid, ApplicantID, ApplicantDept, ref ErrorMsg) == FlowReturn.OK)
            {
                String SerialID = "";
                int Status = 0;
                TmpSql = "SELECT SerialID, Status FROM FSe7en_Sys_Requisition WHERE RequisitionID = @RequisitionID ";
                ParameterList.Clear();
                ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                if (TmpDt.Rows.Count > 0)
                {
                    SerialID = TmpDt.Rows[0]["SerialID"].ToString();
                    Status = Convert.ToInt32(TmpDt.Rows[0]["Status"].ToString());
                }

                Response.Write(OutputJosn(Status, ErrorMsg, RequisitionID));
                /*20200130 修正
                //成功起單後判斷是否為重送表單
                if (OrgRequisitionID != "")
                {
                    TmpSql = " Insert into FSe7en_Tep_ResentList " +
                             "  	  (RequisitionID, OriginalRequisitionID) " +
                             " Select '" + RequisitionID + "', '" + OrgRequisitionID + "' ";
                    ParameterList.Clear();
                    SqlCommand(connString, TmpSql, ParameterList);

                }
                */
            }
            InterfaceConnection.Close();
        }
        return ErrorMsg;
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
    //複製ParameterList
    public static List<SqlParameter> CopyList(List<String> StrList1, List<String> StrList2)
    {
        List<SqlParameter> NewList = new List<SqlParameter>();
        for (int i = 0; i < StrList1.Count; i++)
        {
            NewList.Add(new SqlParameter(StrList1[i], StrList2[i]));
        }
        return NewList;
    }
    //替換特殊字元
    public string ReplaceStr(string FStr)
    {
        FStr = HttpUtility.HtmlEncode(FStr).Replace("&quot;", "\"").Replace("&#39;", "'").Replace("'", "");
        return FStr;
    }
}