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

public partial class ERPCallBPM : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        LogRequest();
        //宣告要ERP起單的表單名稱
        string ErpKey = HttpUtility.HtmlEncode(Request.QueryString["ErpKey"]).Replace("&quot;", "\"").Replace("&#39;", "'") ?? string.Empty;
        string Identify = HttpUtility.HtmlEncode(Request.QueryString["Identify"]) ?? string.Empty;
        string AccountID = HttpUtility.HtmlEncode(Request.QueryString["AccountID"]) ?? string.Empty;
        string DeptID = HttpUtility.HtmlEncode(Request.QueryString["DeptID"]) ?? string.Empty;
        string CounterSignID = HttpUtility.HtmlEncode(Request.QueryString["CounterSignID"]) ?? string.Empty;
        string CounterSignName = HttpUtility.HtmlEncode(Request.QueryString["CounterSignName"]) ?? string.Empty;
        //string ErpKey = "";
        //string Identify = "";
        //string AccountID = "";
        //string DeptID = "";
        //string CounterSignID = "";
        //string CounterSignName = "";
        string OrgApplicantID = "";
        string JsonText = "";
        string TmpSql = "";
        DataTable TmpDt = new DataTable();
        DataTable Dt = new DataTable();

        List<SqlParameter> ParameterList = new List<SqlParameter>();


        //使用AutoWeb資料庫連線字串
        string connString = LoadCmdStr("\\\\Database\\\\Project\\\\SCSB\\\\Flow\\\\connection\\\\BPM.xdbc.xmf", 1);
        JsonText = HttpUtility.HtmlEncode(ErpKey).Replace("&quot;", "\"").Replace("&#39;", "'");
        /*
        //取Json字串
        FileStream fs = new FileStream(Server.MapPath("~/JsonText.txt"), FileMode.Open);
        StreamReader sr = new StreamReader(fs);
        //載入JSON字串
        JsonText = sr.ReadToEnd();
        sr.Close();
        fs.Close();
        Identify = "A61";
        AccountID = "16161";
        DeptID = "011000";
        CounterSignID = "011100";
        CounterSignName = "人力資源處";
        */
        //解JsonToDT
        if (JsonText == "")
        {
            Response.Write(OutputJosn(32, "ErpKey 未輸入", ""));
        }
        else
        {
            JsonText = "[" + ReplaceStr(JsonText) + "]";
            try
            {
                Dt = JsonConvert.DeserializeObject<DataTable>(HttpUtility.HtmlEncode(JsonText).Replace("&quot;", "\"").Replace("&#39;", "'"));

                if (Dt.Columns.Count == 0)
                {
                    Response.Write(OutputJosn(32, "ErpKey資料有誤請檢查", ""));
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                Response.Write(OutputJosn(32, "ErpKey Json格式有誤", ""));
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
                    Response.Write(StarFlow(connString, RequisitionID, Dt, Identify.Replace("'", ""), AccountID.Replace("'", ""), DeptID.Replace("'", ""), OrgApplicantID.Replace("'", ""), CounterSignID.Replace("'", ""), CounterSignName.Replace("'", "")));
                }
            }
        }
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
    //BPM起單
    string StarFlow(string connString, string RequisitionID, DataTable ErpKey, string Identify, string ApplicantID, string ApplicantDept, string OrgApplicantID, string CounterSignID, string CounterSignName)
    {
        string ErrorMsg = "";
        DataTable TmpDt = new DataTable();
        string ITable = Identify;
        string FlowID = Identify;
        string ApplicantName = "";
        string ApplicantDeptName = "";
        string OrgApplicantName = "";
        string DiagramID = "";
        string TmpSql = "";
        string WhereStr = "";
        string WhereStr2 = "";
        string TenderMainSeq = "";
        string AcceptanceSeq = "";
        string Installment = "";
        string BidBondInfoSeq = "";

        string DBName = "";
        int CounterSignCount = 0;
        string OrgRequisitionID = "";

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
        if (!FM7Engine.Initialize("SCSB_Flow", ref ErrorMsg))
        {
            return ErrorMsg;
        }

        //取ERP連線

        if (Identify.IndexOf("A6") != -1 || Identify == "A70" || Identify == "A72" || Identify == "A80")
        {
            TmpSql = "SELECT ConfigValue FROM SCSB_SystemConfig WHERE ConfigID=N'PMS' ";
            ParameterList.Clear();
            TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        }
        else
        {
            TmpSql = "SELECT ConfigValue FROM SCSB_SystemConfig WHERE ConfigID=N'ERPServer' ";
            ParameterList.Clear();
            TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        }

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
        //Identify為A80時，DB的資料皆從A67撈取
        ParameterList.Add(new SqlParameter("@Identify", Identify.Replace("'", "")=="A80" ? "A67" : Identify.Replace("'", "")));
		
        TmpDt = SqlQuery(connString, TmpSql, ParameterList);
        if (TmpDt.Rows.Count > 0)
        {
            DiagramID = TmpDt.Rows[0]["DiagramID"].ToString();
        }
        DiagramGuid sDiagramGuid = new DiagramGuid(DiagramID);

        ParameterList.Clear();
        StrList1.Clear();
        StrList2.Clear();
        //組where條件
        for (int i = 0; i < ErpKey.Columns.Count; i++)
        {
            if (ErpKey.Columns[i].ColumnName == "RequisitionID")
            {
                OrgRequisitionID = ErpKey.Rows[0][i].ToString();
            }
            else
            {
                if (Identify == "A64" || Identify == "A65" || Identify == "A66" || Identify == "A67" || Identify == "A68" || Identify == "A69"|| Identify=="A80")
                {
                    switch (ErpKey.Columns[i].ColumnName)
                    {
                        case "TenderMainSeq":
                            TenderMainSeq = ErpKey.Rows[0][i].ToString();
                            break;
                        case "AcceptanceSeq":
                            AcceptanceSeq = ErpKey.Rows[0][i].ToString();
                            break;
                        case "Installment":
                            Installment = ErpKey.Rows[0][i].ToString();
                            break;
                        case "BidBondInfoSeq":
                            BidBondInfoSeq = ErpKey.Rows[0][i].ToString();
                            break;
                    }
                }
                else
                {
                    if (WhereStr == "")
                    {
                        WhereStr = ErpKey.Columns[i].ColumnName + "=@" + ErpKey.Columns[i].ColumnName;
                        StrList1.Add("@" + ErpKey.Columns[i].ColumnName);
                        StrList2.Add(ErpKey.Rows[0][i].ToString());
                    }
                    else
                    {
                        WhereStr = " and " + ErpKey.Columns[i].ColumnName + "=@" + ErpKey.Columns[i].ColumnName;
                        StrList1.Add("@" + ErpKey.Columns[i].ColumnName);
                        StrList2.Add(ErpKey.Rows[0][i].ToString());
                    }
                }
                if (WhereStr2 == "")
                {
                    if (Identify == "A41" || Identify == "B31" || Identify == "B32" || Identify == "B33" || Identify == "B34")
                    {
                        WhereStr2 = " FormId='" + Identify + "' and OutSeq='" + ErpKey.Rows[0][i].ToString() + "' ";
                    }
                    else
                    {
                        WhereStr2 = " FormId='" + Identify + "' and TenderMainSeq='" + ErpKey.Rows[0][i].ToString() + "' ";
                    }
                }
            }
        }


        //將資料寫入該表單M與D表
        switch (Identify)
        {
            case "A41":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_PlanningBudgetFlow " +
                            "  Where " + WhereStr;

                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);

                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A41， V_PlanningBudgetFlow必須有資料", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //主辦統購單位-
                        TmpSql = " Select distinct DirectUnit, PortalSidUnit, CurrencyName, BudgetLocal " +
                                "   from " + DBName + "dbo.V_PlanningBudgetFlowDetail " +
                                "  Where " + WhereStr;

                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        TmpDt = SqlQuery(connString, TmpSql, ParameterList);


                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName, PlanSeq, PlanNo, " +
                                "        PlanName, PlanYear, SubProcessCount, ExecSubProcess, " +
                                "        DirectUnit, PortalSidUnit, CurrencyName, " +
                                "        PlanFunding, OrgApplicantID, OrgApplicantName) " +
                                " Select distinct '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', PlanSeq, PlanNo,  " +
                                "        PlanName, PlanYear, " + CounterSignCount + " , 0, " +
                                "        '" + TmpDt.Rows[0]["DirectUnit"].ToString() + "','" + TmpDt.Rows[0]["PortalSidUnit"].ToString() + "','" + TmpDt.Rows[0]["CurrencyName"].ToString() + "', " +
                                "        '" + TmpDt.Rows[0]["BudgetLocal"].ToString() + "', '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_PlanningBudgetFlow " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得ERP資料寫入D表
                        TmpSql = " Insert into FM7T_" + Identify + "_D " +
                                 "        (RequisitionID, PlanningBudgetSeq, AccountName, AccountSeq, " +
                                 "  	  OneYearBudget,LastYearTotalBudget,BudgetLocal, LocalRate, " +
                                 " 	      TotalBudget, Memo, BudgetYear) " +
                                 " Select '" + RequisitionID + "', PlanningBudgetSeq, AccountName, AccountSeq, " +
                                 " 	      OneYearBudget,LastYearTotalBudget,BudgetLocal, LocalRate, " +
                                 " 	      TotalBudget, Memo, BudgetYear " +
                                 "   from " + DBName + "dbo.V_PlanningBudgetFlowDetail " +
                                 "  where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得預算數編列附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by Seq), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Budget " +
                                 "  where " + WhereStr2 +
                                 "  order by Seq ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    return OutputJosn(99, "SQL發生錯誤請檢查SQL是否正確", "");
                }
                break;
            case "B31":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_PlanningBudgetAbsorp " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);

                    if (TmpDt.Rows.Count == 0)
                    {
                        return OutputJosn(99, "Identify＝B31，V_PlanningBudgetAbsorp必須有資料", "");
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, PlanSeq, PlanNo, " +
                                "        PlanName, PlanYear, PlanBudgetAbsorpAuditSeq, DirectUnitName, PurcharseUnitName, " +
                                "        PlanKindName, PlanFunding, AuditMemo, OrgApplicantID, OrgApplicantName) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, PlanSeq, PlanNo,  " +
                                "        PlanName, PlanYear, PlanBudgetAbsorpAuditSeq , DirectUnitName, PurcharseUnitName, " +
                                "        PlanKindName, PlanFunding, AuditMemo, '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_PlanningBudgetAbsorp " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);


                        //取得ERP資料寫入D表
                        TmpSql = " Insert into FM7T_" + Identify + "_D " +
                                 "        (RequisitionID, PlanBudgetAbsorpAuditSeq, CurrencyName, AccountName, " +
                                 "  	  AccountSeq, BudgetLocal, AbsortLocal, BudgetNT, AbsorpNT, " +
                                 " 	      LocalRate, Memo, TotalBudget) " +
                                 " Select '" + RequisitionID + "', PlanBudgetAbsorpAuditSeq, CurrencyName, AccountName, " +
                                 " 	      AccountSeq, BudgetLocal, AbsortLocal, BudgetNT, AbsorpNT, " +
                                 " 	      LocalRate, Memo, TotalBudget " +
                                 "   from " + DBName + "dbo.V_PlanningBudgetAbsorpDetail " +
                                 "  where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得勻支附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by Seq), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Budget " +
                                 "  where " + WhereStr2 +
                                 "  order by Seq ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    Response.Write(OutputJosn(99, "SQL發生錯誤請檢查SQL是否正確", ""));
                }
                break;
            case "B32":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_BTPlanBudgetFlowDetail " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        return OutputJosn(99, "Identify＝B32，V_BTPlanBudgetFlowDetail必須有資料", "");
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag,  PlanBudgetFlowSeq, PlanYear,  " +
                                "        OutPlanNo, OutPlanName, OutAccountName, OutPlanMainDate, OutCurrencyName, " +
                                "        OutBudgetTotalBudget, OutBudgetPay, OutLastBudget, OutLocal, InPlanNo, " +
                                "        InPlanName, InAccountName, InPlanMainDate, InCurrencyName, InBudgetTotalBudget, " +
                                "        InBudgetPay, InLastBudget, InLocal, FlowMemo, OrgApplicantID, OrgApplicantName) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, PlanBudgetFlowSeq, PlanYear, " +
                                "        OutPlanNo, OutPlanName, OutAccountName, OutPlanMainDate, OutCurrencyName, " +
                                "        OutBudgetTotalBudget, OutBudgetPay, OutLastBudget, OutLocal, InPlanNo, " +
                                "        InPlanName, InAccountName, InPlanMainDate, CurrencyName, InBudgetTotalBudget, " +
                                "        InBudgetPay, InLastBudget, InLocal, FlowMemo, '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_BTPlanBudgetFlowDetail " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得流用附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by Seq), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Budget " +
                                 "  where " + WhereStr2 +
                                 "  order by Seq ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, "SQL發生錯誤請檢查SQL是否正確", "");
                    return ErrorMsg;
                }
                break;
            case "B33":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_BTPlanBudgetAdjustDetail " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝B33，V_BTPlanBudgetAdjustDetail必須有資料", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName,  " +
                                "        SubProcessCount, ExecSubProcess, PlanBudgetAdjustSeq, PlanNo, PlanName, " +
                                "        PlanKindName, PlanYear, PlanFunding, PlanMainDate, AccountName, " +
                                "        AdjType, AdjustLocal, AdjustMemo, AuditStatus, OrgApplicantID, OrgApplicantName) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', " +
                                "        " + CounterSignCount + ", 0, PlanBudgetAdjustSeq, PlanNo, PlanName, " +
                                "        PlanKindName, PlanYear, PlanFunding, PlanMainDate, AccountName,  " +
                                "        AdjType, AdjustLocal, AdjustMemo, AuditStatus, '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_BTPlanBudgetAdjustDetail " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得追加附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by Seq), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Budget " +
                                 "  where " + WhereStr2 +
                                 "  order by Seq ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, "SQL發生錯誤請檢查SQL是否正確", "");
                    return ErrorMsg;
                }
                break;
            case "B34":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_BTPlanBudgetAdjustDetail " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝B34，V_BTPlanBudgetAdjustDetail必須有資料", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, PlanNo, PlanName, " +
                                "        PlanBudgetAdjustSeq, PlanKindName, PlanYear, PlanFunding, PlanMainDate, AccountName, " +
                                "        AdjType, AdjustLocal, AdjustMemo, AuditStatus, OrgApplicantID, OrgApplicantName) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, PlanNo, PlanName,  " +
                                "        PlanBudgetAdjustSeq, PlanKindName, PlanYear, PlanFunding, PlanMainDate, AccountName,  " +
                                "        AdjType, AdjustLocal, AdjustMemo, AuditStatus, '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_BTPlanBudgetAdjustDetail " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得追減附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by Seq), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Budget " +
                                 "  where " + WhereStr2 +
                                 "  order by Seq ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, "SQL發生錯誤請檢查SQL是否正確", "");
                    return ErrorMsg;
                }
                break;
            case "A61":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_GRMain_Export " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A61，V_GRMain_Export必須有資料，一般請款", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        

                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, OrgApplicantID, OrgApplicantName, GRMainSeq, GRYear,  " +
                                "        ReceiveDate, GRMainNo, GRMainName, UnitName, UnitSeq, " +
                                "        Replace, Reason, GRMainType, TotalPrice ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + OrgApplicantID + "','" + OrgApplicantName + "', GRMainSeq, GRYear,  " +
                                "        ReceiveDate, GRMainNo, GRMainName, UnitName, UnitSeq, " +
                                "        Replace, Reason, GRMainType, TotalPrice " +
                                "   from " + DBName + "dbo.V_GRMain_Export " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);



                        //取得預算來源寫入D表
                        TmpSql = " Insert into FM7T_" + Identify + "_D " +
                                 "        (RequisitionID, GRWorkItemSeq, GRMainSeq, PlanYear, PlanName, " +
                                 "  	  AccountName, LeftBudget, ContractorName, BankAccount, IsChargeInside, " +
                                 "        InvoiceType, InvoiceNo, InvoiceDate, InvoiceCurrency, InvoiceAmount, " +
                                 "        ExchangeRate, CurrencyCode, ItemPrice, WriteOffMemo, CreditCardPayment, " +
                                 "        ContractorTaxId, BankName, BankAccountName, AccountingStaffID, "+
                                 "        DepositAccount, CreditCardAccount,InvoiceCurrencyCode ) " +
                                 " Select '" + RequisitionID + "', GRWorkItemSeq, GRMainSeq, PlanYear, PlanName, " +
                                 "  	  AccountName, LeftBudget, ContractorName, BankAccount, IsChargeInside, " +
                                 "        InvoiceType, InvoiceNo, InvoiceDate, InvoiceCurrency, InvoiceAmount, " +
                                 "        ExchangeRate, CurrencyCode, ItemPrice, WriteOffMemo, CreditCardPayment, " +
                                 "        ContractorTaxId, BankName, BankAccountName, UserName, " +
                                 "        DepositAccount, CreditCardAccount,InvoiceCurrencyCode " +
                                 "   from " + DBName + "dbo.V_GRWorkItem_Export " +
                                 "  where " + WhereStr +
                                 "  order by GRWorkItemSeq ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_GRFile_Export " +
                                 "  where " + WhereStr +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

 /*   //A61 無CaseNo           
                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
 */
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A62":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_TenderMain0_Export " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A62，V_TenderMain0_Export必須有資料，請購擬稿作業", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, TenderMainSeq, TenderStageSeq, " +
                                "        CaseNo, CaseName, UnitInfo, UnitID, UnitName, UserInfo, " +
                                "        UserName, EmployeeID, IsDesManu, IsStakeholder, OrgApplicantID, OrgApplicantName, " +
                                "        ReasonDemand, CurrencyCode, Budget, TotalPrice) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', " + CounterSignCount + ", 0, TenderMainSeq, TenderStageSeq,  " +
                                "        CaseNo, CaseName, UnitInfo, UnitID, UnitName, UserInfo,  " +
                                "        UserName, EmployeeID, IsDesManu, IsStakeholder, '" + OrgApplicantID + "','" + OrgApplicantName + "', " +
                                "        ReasonDemand, CurrencyCode, Budget, TotalPrice  " +
                                "   from " + DBName + "dbo.V_TenderMain0_Export " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                        

                        //取得關係購案寫入D表
                        TmpSql = " Insert into FM7T_" + Identify + "_D " +
                                 "        (RequisitionID, TenderMainIncludeSeq, TenderMainSeq, CaseNo, " +
                                 "  	  RelCaseName, OrderNo) " +
                                 " Select '" + RequisitionID + "', TenderMainIncludeSeq, TenderMainSeq, CaseNo, " +
                                 " 	       RelCaseName, OrderNo " +
                                 "   from " + DBName + "dbo.V_TenderMainInclude_Export " +
                                 "  where " + WhereStr +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得預算來源資料
                        TmpSql = " Insert into FM7T_A62_TenderBSWorkItem_Export " +
                                 "        (RequisitionID, TenderBSWorkItemSeq, TenderMainSeq, BudgetYear, PlanName, " +
                                 "  	  AccountName, TaskSubject, Unit, Quantity, UnitPrice, " +
                                 "        CurrencyCode, TotalUnitPrice, TotalUnitPriceCurrencyCode, Memo) " +
                                 " Select '" + RequisitionID + "', TenderBSWorkItemSeq, TenderMainSeq, BudgetYear, PlanName, " +
                                 " 	       AccountName, TaskSubject, Unit, Quantity, UnitPrice, " +
                                 " 	       CurrencyCode, TotalUnitPrice, TotalUnitPriceCurrencyCode, Memo " +
                                 "   from " + DBName + "dbo.V_TenderBSWorkItem_Export " +
                                 "  where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  where " + WhereStr + " AND AcceptanceSeq='15' "+ //20200727 新規格
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        
                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A63":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_TenderMain1_Export " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A63，V_TenderMain1_Export必須有資料，採購擬稿作業", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, TenderMainSeq, TenderStageSeq, " +
                                "        CaseNo, CaseName, UnitInfo, UnitID, UnitName, UserInfo, " +
                                "        UserName, EmployeeID, IsDesManu, IsStakeholder, OrgApplicantID, OrgApplicantName, " +
                                "        ReasonDemand, CurrencyCode, Budget, TotalPrice, TenderCOMM) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', " + CounterSignCount + ", 0, TenderMainSeq, TenderStageSeq,  " +
                                "        CaseNo, CaseName, UnitInfo, UnitID, UnitName, UserInfo,  " +
                                "        UserName, EmployeeID, IsDesManu, IsStakeholder, '" + OrgApplicantID + "','" + OrgApplicantName + "', " +
                                "        ReasonDemand, CurrencyCode, Budget, TotalPrice, TenderCOMM  " +
                                "   from " + DBName + "dbo.V_TenderMain1_Export " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);
                        //取得關係購案寫入D表
                        TmpSql = " Insert into FM7T_" + Identify + "_D " +
                                 "        (RequisitionID, TenderMainIncludeSeq, TenderMainSeq, CaseNo, " +
                                 "  	  RelCaseName, OrderNo) " +
                                 " Select '" + RequisitionID + "', TenderMainIncludeSeq, TenderMainSeq, CaseNo, " +
                                 " 	       RelCaseName, OrderNo " +
                                 "   from " + DBName + "dbo.V_TenderMainInclude_Export " +
                                 "  where " + WhereStr +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得預算來源資料
                        TmpSql = " Insert into FM7T_A63_TenderBSWorkItem_Export " +
                                 "        (RequisitionID, TenderBSWorkItemSeq, TenderMainSeq, BudgetYear, PlanName, " +
                                 "  	  AccountName, TaskSubject, Unit, Quantity, UnitPrice, " +
                                 "        CurrencyCode, TotalUnitPrice, TotalUnitPriceCurrencyCode, Memo) " +
                                 " Select '" + RequisitionID + "', TenderBSWorkItemSeq, TenderMainSeq, BudgetYear, PlanName, " +
                                 " 	       AccountName, TaskSubject, Unit, Quantity, UnitPrice, " +
                                 " 	       CurrencyCode, TotalUnitPrice, TotalUnitPriceCurrencyCode, Memo " +
                                 "   from " + DBName + "dbo.V_TenderBSWorkItem_Export " +
                                 "  where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  where " + WhereStr + " AND  AcceptanceSeq='1' "+  //20200727 新規格
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A64":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_ContractOperation_Export " +
                            "  Where TenderMainSeq=@TenderMainSeq ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A64，V_ContractOperation_Export必須有資料，比議價/決標作業", "");
                        return ErrorMsg;
                    }
                    else
                    {            
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, TenderMainSeq, TenderStageSeq, " +
                                "        CaseNo, CaseName, UnitInfo, UnitID, UnitName, UserInfo, " +
                                "        UserName, EmployeeID, TotalDecidePriceCurrencyCode, TotalDecidePrice, OrgApplicantID, OrgApplicantName, " +
                                "        Process, CurrencyCode, CurrencyName, Budget, PayScheduleMemo, IsContractSchedule, EstimatedAmount, DecidePrice,TotalPrice) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', '" + CounterSignCount + "', 0, TenderMainSeq, TenderStageSeq,  " +
                                "        CaseNo, CaseName, UnitInfo, UnitID, UnitName, UserInfo,  " +
                                "        UserName, EmployeeID, TotalDecidePriceCurrencyCode, TotalDecidePrice, '" + OrgApplicantID + "','" + OrgApplicantName + "', " +
                                "        Process, CurrencyCode, CurrencyName, Budget, PayScheduleMemo, IsContractSchedule, EstimatedAmount, DecidePrice,TotalPrice  " +
                                "   from " + DBName + "dbo.V_ContractOperation_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";
                        /*
                        TmpSql +=@"
                                    INSERT INTO [dbo].[PMS_V_ContractOperation_Export]
                                    SELECT '"+RequisitionID+@"', *, getdate()  FROM "+DBName+@"dbo.V_ContractOperation_Export
                        ";
                        */

                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);
                        
                        //取得投標廠商
                        TmpSql = " Insert into FM7T_A64_All_Contractor_Expor " +
                                 "        (RequisitionID, TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 "  	  CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid, "+
								 "         CheckResult, RelatedPerson,DecidePrice) " +
                                 " Select '" + RequisitionID + "', Contractor.TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 " 	       CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid, " +
								 "         CASE WHEN ISNULL(Stakeholder.RAE_IDNO,'')<>'' THEN 'Y' ELSE 'N' END AS 'CheckResult' , Stakeholder.DiplayName,DecidePrice "+
                                 "   from " + DBName + "dbo.V_All_Contractor_Export Contractor left outer join "+ DBName + "dbo.V_Stakeholder_Export Stakeholder "+
								 "     on Contractor.TaxId =Stakeholder.RAE_IDNO and Contractor.TenderMainSeq =Stakeholder.TenderMainSeq "+
                                 "  Where Contractor.TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得採購明細
                        TmpSql = " Insert into FM7T_A64_All_PerformBSWorkItem_Export " +
                                 "        (RequisitionID, TenderMainSeq, BudgetYear, PlanName, AccountName, " +
                                 "  	  TaskSubject, Unit, Quantity, CurrencyCode, CurrencyName, " +
                                 "        UnitPrice, TotalPrice, Memo) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, BudgetYear, PlanName, AccountName, " +
                                 " 	       TaskSubject, Unit, Quantity, CurrencyCode, CurrencyName, " +
                                 " 	       UnitPrice, TotalPrice, Memo " +
                                 "   from " + DBName + "dbo.V_All_PerformBSWorkItem_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得履約、保固保證金
                        TmpSql = " Insert into FM7T_A64_BidBondInfo_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	       BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 " 	       CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金
                        TmpSql = " Insert into FM7T_A64_BidBondInfo_History_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and TaxId=(SELECT top 1 TaxId FROM " + DBName + "dbo.V_Contractor_Export where IsBid='1' and TenderMainSeq=@TenderMainSeq_2 )";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@TenderMainSeq_2", TenderMainSeq.Replace("'", "")));
                      //  ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type, UseSeal) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1, UseSeal " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and FormId='A64' " + " AND  AcceptanceSeq>=3 AND AcceptanceSeq<>15  " +//20200727 新規格
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);   
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A65":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_AcceptanceOperation_Export " +
                            "  Where TenderMainSeq=@TenderMainSeq ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A65，V_AcceptanceOperation_Export必須有資料，驗收作業", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表 
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, " +
                                "        TenderMainSeq, TenderStageSeq, CaseNo, CaseName, " +
                                "        UnitInfo, UnitID, UnitName, UserInfo, UserName, " +
                                "        EmployeeID, TotalDecidePriceCurrencyCode, TotalDecidePrice, CurrentAmount, PreviousAmount, " +
                                "        PreviousVariationAmount, OrgApplicantID, OrgApplicantName ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', " + CounterSignCount + ", 0, " +
                                "        TenderMainSeq, TenderStageSeq, CaseNo, CaseName, " +
                                "        UnitInfo, UnitID, UnitName, UserInfo, UserName,  " +
                                "        EmployeeID, TotalDecidePriceCurrencyCode, TotalDecidePrice, CurrentAmount, PreviousAmount, " +
                                "        PreviousVariationAmount, '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_AcceptanceOperation_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";

                        TmpSql = TmpSql + "INSERT INTO [dbo].[PMS_V_AcceptanceOperation_Export] "+
                                  " SELECT '"+RequisitionID+@"',* ,getdate()  FROM "+DBName+@"dbo.V_AcceptanceOperation_Export";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //本次驗收紀錄
                        TmpSql = " Insert into FM7T_A65_Current_All_Acceptance_Export " +
                                 "        (RequisitionID, TenderMainSeq, AcceptanceSeq, Installment, ModeName, " +
                                 "  	  CurrentDate, CurrencyCode, CurrencyName, Amount, LiquidatedDamages, " +
                                 "        AcceptanceChargeback, Result, Memo, IsPaid, TotalPrice) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, AcceptanceSeq, Installment, ModeName, " +
                                 " 	       Date, CurrencyCode, CurrencyName, Amount, LiquidatedDamages, " +
                                 " 	       AcceptanceChargeback, Result, Memo, IsPaid, TotalPrice " +
                                 "   from " + DBName + "dbo.V_Current_All_Acceptance_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and AcceptanceSeq=@AcceptanceSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@AcceptanceSeq", AcceptanceSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史驗收紀錄
                        TmpSql = " Insert into FM7T_A65_History_All_Acceptance_Export " +
                                 "        (RequisitionID, TenderMainSeq, AcceptanceSeq, Installment, ModeName, " +
                                 "  	  CurrentDate, CurrencyCode, CurrencyName, Amount, LiquidatedDamages, " +
                                 "        AcceptanceChargeback, Result, Memo, IsPaid, TotalPrice) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, AcceptanceSeq, Installment, ModeName, " +
                                 " 	       Date, CurrencyCode, CurrencyName, Amount, LiquidatedDamages, " +
                                 " 	       AcceptanceChargeback, Result, Memo, IsPaid, TotalPrice " +
                                 "   from " + DBName + "dbo.V_History_All_Acceptance_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";

                        TmpSql += @"
                                    INSERT INTO [dbo].[PMS_V_History_All_Acceptance_Export]
                                    SELECT '"+RequisitionID+@"',*,getdate()  FROM "+DBName+@"dbo.V_History_All_Acceptance_Export 
                            ";         
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //檢視驗收內容
                        TmpSql = " Insert into FM7T_A65_All_AcceptanceBSWorkItem_Export " +
                                 "        (RequisitionID, TenderMainSeq, AcceptanceSeq, BudgetYear, PlanName, " +
                                 "  	  AccountName, TaskSubject, Unit, Quantity, QuantityLeft, " +
                                 "        ExcutionQuantity, UnitPrice, TotalPrice, Memo) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, AcceptanceSeq, BudgetYear, PlanName, " +
                                 " 	       AccountName, TaskSubject, Unit, Quantity, QuantityLeft, " +
                                 " 	       ExcutionQuantity, UnitPrice, TotalPrice, Memo " +
                                 "   from " + DBName + "dbo.V_All_AcceptanceBSWorkItem_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and AcceptanceSeq=@AcceptanceSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@AcceptanceSeq", AcceptanceSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and AcceptanceSeq=@AcceptanceSeq and FormId='A65' " +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@AcceptanceSeq", AcceptanceSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A66":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_AcceptanceOperation_Export " +
                            "  Where TenderMainSeq=@TenderMainSeq ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A66，V_AcceptanceOperation_Export必須有資料，驗收作業-付款", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, TenderMainSeq, TenderStageSeq, CaseNo, CaseName, " +
                                "        UnitInfo, UnitID, UnitName, UserInfo, UserName, " +
                                "        EmployeeID, TotalDecidePriceCurrencyCode, TotalDecidePrice, CurrentAmount, PreviousAmount, " +
                                "        OrgApplicantID, OrgApplicantName ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, TenderMainSeq, TenderStageSeq, CaseNo, CaseName, " +
                                "        UnitInfo, UnitID, UnitName, UserInfo, UserName,  " +
                                "        EmployeeID, TotalDecidePriceCurrencyCode, TotalDecidePrice, CurrentAmount, PreviousAmount, " +
                                "        '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_AcceptanceOperation_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //本次驗收紀錄
                        TmpSql = " Insert into FM7T_A66_Current_All_Acceptance_Export " +
                                 "        (RequisitionID, TenderMainSeq, AcceptanceSeq, Installment, ModeName, " +
                                 "  	  CurrentDate, CurrencyCode, CurrencyName, Amount, LiquidatedDamages, " +
                                 "        AcceptanceChargeback, Result, Memo, IsPaid, TotalPrice, CaseType) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, AcceptanceSeq, Installment, ModeName, " +
                                 " 	       Date, CurrencyCode, CurrencyName, Amount, LiquidatedDamages, " +
                                 " 	       AcceptanceChargeback, Result, Memo, IsPaid, TotalPrice, CaseType " +
                                 "   from " + DBName + "dbo.V_Current_All_Acceptance_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //本次付款
                        TmpSql = " Insert into FM7T_A66_Current_All_PayRecord_Export " +
                                 "        (RequisitionID, TenderMainSeq, AcceptanceSeq, Installment, Date, " +
                                 "  	  PayAmount, InvoiceNumber,TaxId,BankAccount,BankMajor,BankTitle,Name) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, AcceptanceSeq, Installment, Date, " +
                                 " 	      PayAmount, InvoiceNumber,TaxId,BankAccount,BankMajor,BankTitle,Name" +
                                 "   from " + DBName + "dbo.V_Current_All_PayRecord_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and AcceptanceSeq=@AcceptanceSeq and Installment=@Installment ";

                        TmpSql +=@"
                                INSERT INTO [dbo].[PMS_V_Current_All_PayRecord_Export]
                                SELECT '"+RequisitionID+@"',*,getdate()  FROM "+DBName+@"dbo.V_Current_All_PayRecord_Export
                        ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@AcceptanceSeq", AcceptanceSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@Installment", Installment.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //檢視驗收內容
                        TmpSql = " Insert into FM7T_A66_All_AcceptanceBSWorkItem_Export " +
                                 "        (RequisitionID, TenderMainSeq, AcceptanceSeq, BudgetYear, PlanName, " +
                                 "  	  AccountName, TaskSubject, Unit, Quantity, QuantityLeft, " +
                                 "        ExcutionQuantity, UnitPrice, CurrencyCode, CurrencyName, TotalPrice, Memo) " +
                                 " Select '" + RequisitionID + "',  TenderMainSeq, AcceptanceSeq, BudgetYear, PlanName, " +
                                 " 	       AccountName, TaskSubject, Unit, Quantity, QuantityLeft, " +
                                 " 	       ExcutionQuantity, UnitPrice, CurrencyCode, CurrencyName, TotalPrice, Memo " +
                                 "   from " + DBName + "dbo.V_All_AcceptanceBSWorkItem_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and AcceptanceSeq=@AcceptanceSeq and FormId='A66' " +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@AcceptanceSeq", AcceptanceSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A67":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_PerformContract_Export " +
                            "  Where TenderMainSeq=@TenderMainSeq ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A67，V_PerformContract_Export必須有資料，履約作業", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                "        CurrencyName, PURCAmount, OrgApplicantID, OrgApplicantName, PerformContractSeq ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                "        CurrencyName, PURCAmount, '" + OrgApplicantID + "','" + OrgApplicantName + "', PerformContractSeq " +
                                "   from " + DBName + "dbo.V_PerformContract_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //履約廠商
                        TmpSql = " Insert into FM7T_A67_All_Bid_Contractor_Export " +
                                 "        (RequisitionID, TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 "  	  CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid ) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 " 	       CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid " +
                                 "   from " + DBName + "dbo.V_All_Bid_Contractor_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //履約、保固保證金
                        TmpSql = " Insert into FM7T_A67_BidBondInfo_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金
                        TmpSql = " Insert into FM7T_A67_BidBondInfo_History_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq<>@BidBondInfoSeq   and TaxId=( select top 1 TaxId  from   " + DBName + "dbo.V_All_Bid_Contractor_Export Where TenderMainSeq=@TenderMainSeq_3  )   ";
                         /*   TmpSql+=@" 
                                    INSERT INTO [dbo].[PMS_V_BidBondInfo_History_Export]
                                    SELECT '"+RequisitionID+@"' ,*,getdate()  FROM "+DBName+@"dbo.V_BidBondInfo_History_Export
                        ";
                        */
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@TenderMainSeq_3", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and FormId='A67'  " +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

						//履約、保固保證金附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
								 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);
                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A68":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_PerformContract_Export " +
                            "  Where TenderMainSeq=@TenderMainSeq ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A68，V_PerformContract_Export必須有資料，履約作業退回", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, " +
                                "        DraftFlag, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                "        CurrencyName, PURCAmount, OrgApplicantID, OrgApplicantName, PerformContractSeq ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        '" + CounterSignID + "','" + CounterSignName + "', " + CounterSignCount + ", 0, " +
                                "        0, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                "        CurrencyName, PURCAmount, '" + OrgApplicantID + "','" + OrgApplicantName + "', PerformContractSeq " +
                                "   from " + DBName + "dbo.V_PerformContract_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //履約廠商
                        TmpSql = " Insert into FM7T_A68_All_Bid_Contractor_Export " +
                                 "        (RequisitionID, TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 "  	  CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid ) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 " 	       CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid " +
                                 "   from " + DBName + "dbo.V_All_Bid_Contractor_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //履約、保固保證金
                        TmpSql = " Insert into FM7T_A68_BidBondInfo_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金
                        TmpSql = " Insert into FM7T_A68_BidBondInfo_History_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq<>@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and FormId='A68' " +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);
						
			//履約、保固保證金附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
								 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);
						
                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A69":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_PerformContract_Export " +
                            "  Where TenderMainSeq=@TenderMainSeq ";
                    ParameterList.Clear();
                    ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A69，V_PerformContract_Export必須有資料，履約轉保固", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, " +
                                "        DraftFlag, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                "        CurrencyName, PURCAmount, OrgApplicantID, OrgApplicantName, PerformContractSeq ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        '" + CounterSignID + "','" + CounterSignName + "', " + CounterSignCount + ", 0, " +
                                "        0, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                "        CurrencyName, PURCAmount, '" + OrgApplicantID + "','" + OrgApplicantName + "', PerformContractSeq " +
                                "   from " + DBName + "dbo.V_PerformContract_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //履約廠商
                        TmpSql = " Insert into FM7T_A69_All_Bid_Contractor_Export " +
                                 "        (RequisitionID, TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 "  	  CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid ) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                 " 	       CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid " +
                                 "   from " + DBName + "dbo.V_All_Bid_Contractor_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //履約、保固保證金
                        TmpSql = " Insert into FM7T_A69_BidBondInfo_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金
                        TmpSql = " Insert into FM7T_A69_BidBondInfo_History_Export " +
                                 "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate) " +
                                 " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                 " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                 "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                 "        ForecastMaturityDate " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and FormId='A69'  " +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);
						
						//履約、保固保證金附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史履約、保固保證金附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
								 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                        ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                        SqlCommand(connString, TmpSql, ParameterList);
						
                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A70":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_BidAnnouncement_Export " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A70，V_BidAnnouncement_Export必須有資料，招標公告", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, TenderMainSeq, TenderStageSeq, CurrencyCode, CurrencyName, " +
                                "        CaseNo, CaseName, SubmitDate, CloseDate, " +
                                "        OpeningTimeS, OpeningPlace, OrgApplicantID, OrgApplicantName, Memo) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, TenderMainSeq, TenderStageSeq, CurrencyCode, CurrencyName, " +
                                "        CaseNo, CaseName, SubmitDate, CloseDate,  " +
                                "        OpeningTimeS, OpeningPlace, '" + OrgApplicantID + "','" + OrgApplicantName + "', Memo " +
                                "   from " + DBName + "dbo.V_BidAnnouncement_Export " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  where " + WhereStr + "  AND AcceptanceSeq='2'  "+  //20200727 新規格
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);



                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
            case "A72":
                try
                {
                    //判斷View是否有值
                    TmpSql = " Select * " +
                            "   from " + DBName + "dbo.V_AcceptanceOperation_Export " +
                            "  Where " + WhereStr;
                    ParameterList.Clear();
                    ParameterList = CopyList(StrList1, StrList2);
                    TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                    if (TmpDt.Rows.Count == 0)
                    {
                        ErrorMsg = OutputJosn(99, "Identify＝A72，V_AcceptanceOperation_Export必須有資料，訂交貨管理", "");
                        return ErrorMsg;
                    }
                    else
                    {
                        //取得ERP資料寫入M表
                        TmpSql = " Insert into FM7T_" + Identify + "_M " +
                                "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                "        DraftFlag, CounterSignID, CounterSignName, SubProcessCount, ExecSubProcess, " +
                                "        TenderMainSeq, TenderStageSeq, CaseNo, CaseName, " +
                                "        UnitInfo, UnitID, UnitName, UserInfo, UserName, " +
                                "        EmployeeID, TotalDecidePriceCurrencyCode,  " +
                                "        OrgApplicantID, OrgApplicantName ) " +
                                " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                "        0, '" + CounterSignID + "','" + CounterSignName + "', " + CounterSignCount + ", 0, " +
                                "        TenderMainSeq, TenderStageSeq, CaseNo, CaseName, " +
                                "        UnitInfo, UnitID, UnitName, UserInfo, UserName,  " +
                                "        EmployeeID, TotalDecidePriceCurrencyCode, " +
                                "        '" + OrgApplicantID + "','" + OrgApplicantName + "' " +
                                "   from " + DBName + "dbo.V_AcceptanceOperation_Export " +
                                "  Where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //本次訂交貨紀錄
                        TmpSql = " Insert into FM7T_A72_Current_AcceptanceOperationBSWorkItem " +
                                 "        (RequisitionID, TenderMainSeq, CaseNo, AcceptanceSeq,  AcceptanceBSWorkItemSeq, " +
                                 "  	  BudgetYear, PlanNo, PlanName, AccountName, TaskSubject, " +
                                 "        Unit, Quantity, UnitPrice, ItemPrice, CurrencyCode, " +
                                 "        CurrencyName, VariationQuantity, VariationAmount, SchTranAmount, Memo, " +
                                 "        PropertyNo, LeftBudget) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, CaseNo, AcceptanceSeq,  AcceptanceBSWorkItemSeq, " +
                                 "  	  BudgetYear, PlanNo, PlanName, AccountName, TaskSubject, " +
                                 "        Unit, Quantity, UnitPrice, ItemPrice, CurrencyCode, " +
                                 "        CurrencyName, VariationQuantity, VariationAmount, SchTranAmount, Memo, " +
                                 "        PropertyNo, LeftBudget " +
                                 "   from " + DBName + "dbo.V_AcceptanceOperationBSWorkItem " +
                                 "  where " + WhereStr;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //歷史訂交貨紀錄
                        WhereStr2 = " CaseNo='" + TmpDt.Rows[0]["CaseNo"].ToString() + "' And RequisitionID in (Select RequisitionID from FSe7en_Sys_Requisition where left(DiagramID,3)='A72' and Status=1) ";

                        TmpSql = " Insert into FM7T_A72_History_AcceptanceOperationBSWorkItem " +
                                 "        (RequisitionID, TenderMainSeq, CaseNo, AcceptanceSeq,  AcceptanceBSWorkItemSeq, " +
                                 "  	  BudgetYear, PlanNo, PlanName, AccountName, TaskSubject, " +
                                 "        Unit, Quantity, UnitPrice, ItemPrice, CurrencyCode, " +
                                 "        CurrencyName, VariationQuantity, VariationAmount, SchTranAmount, Memo, " +
                                 "        PropertyNo, LeftBudget) " +
                                 " Select '" + RequisitionID + "', TenderMainSeq, CaseNo, AcceptanceSeq,  AcceptanceBSWorkItemSeq, " +
                                 "  	  BudgetYear, PlanNo, PlanName, AccountName, TaskSubject, " +
                                 "        Unit, Quantity, UnitPrice, ItemPrice, CurrencyCode, " +
                                 "        CurrencyName, VariationQuantity, VariationAmount, SchTranAmount, Memo, " +
                                 "        PropertyNo, LeftBudget " +
                                 "   from dbo.FM7T_A72_Current_AcceptanceOperationBSWorkItem " +
                                 "  where " + WhereStr2;
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //取得購案附件
                        TmpSql = " Insert into FM7T_" + Identify + "_F " +
                                 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                 " Select (Select isnull(max(AutoCounter),0) from FM7T_" + Identify + "_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                 " 	       FullPath, FileName, 0, 0, 1 " +
                                 "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                 "  where " + WhereStr +
                                 "  order by OrderNo ";
                        ParameterList.Clear();
                        ParameterList = CopyList(StrList1, StrList2);
                        SqlCommand(connString, TmpSql, ParameterList);

                        //關聯表單
                        TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@Identify", Identify));
                        ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                        SqlCommand(connString, TmpSql, ParameterList);
                    }
                }
                catch (Exception ex)
                {
                    ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                    return ErrorMsg;
                }
                break;
                case "A80":
                    try
                    {
                        //判斷View是否有值
                        TmpSql = " Select * " +
                                "   from " + DBName + "dbo.V_TenderMain3_Export " +
                                "  Where TenderMainSeq=@TenderMainSeq ";
                        ParameterList.Clear();
                        ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq));
                        TmpDt = SqlQuery(connString, TmpSql, ParameterList);
                        if (TmpDt.Rows.Count == 0)
                        {
                            ErrorMsg = OutputJosn(99, "Identify＝A80，V_TenderMain3_Export必須有資料，開標及審標作業", "");
                            return ErrorMsg;
                        }
                        else
                        {
                            //取得ERP資料寫入M表
                            TmpSql = " Insert into FM7T_A67_M " +
                                    "       (RequisitionID,DiagramID,ApplicantDept,ApplicantDeptName,ApplicantID, " +
                                    "        ApplicantName,FillerID,FillerName,ApplicantDateTime,Priority, " +
                                    "        DraftFlag, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                    "        CurrencyName, PURCAmount, OrgApplicantID, OrgApplicantName, PerformContractSeq ) " +
                                    " Select '" + RequisitionID + "','" + DiagramID + "','" + ApplicantDept + "','" + ApplicantDeptName + "','" + ApplicantID + "', " +
                                    "        '" + ApplicantName + "','" + ApplicantID + "','" + ApplicantName + "', GetDate(), 2, " +
                                    "        0, TenderMainSeq, CaseNo, CaseName, CurrencyCode, " +
                                    "        '', TotalPrice, '" + OrgApplicantID + "','" + OrgApplicantName + "', '' " +
                                    "   from " + DBName + "dbo.V_TenderMain3_Export " +
                                    "  Where TenderMainSeq=@TenderMainSeq ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);

                            //履約廠商==>issue 20200609
                            TmpSql = " Insert into FM7T_A67_All_Bid_Contractor_Export " +
                                    "        (RequisitionID, TenderMainSeq, Contractor, TaxId, CurrencyCode, " +
                                    "  	  CurrencyName, Price, IsParticipateSelection, IsParticipateBargain, IsBid ) " +
                                    " Select '" + RequisitionID + "', TenderMainSeq, Contractor, TaxId, '.', " +
                                    " 	       '.', Price, '0', '0', '1' " +
                                    "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                    "  Where TenderMainSeq=@TenderMainSeq ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);



                        
                            //履約、保固保證金
                            TmpSql = " Insert into FM7T_A67_BidBondInfo_Export " +
                                     "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                     "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                     "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                     "        ForecastMaturityDate) " +
                                     " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                     " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                     "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                     "        ForecastMaturityDate " +
                                     "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                                     "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);

                            //歷史履約、保固保證金
                            TmpSql = " Insert into FM7T_A67_BidBondInfo_History_Export " +
                                     "        (RequisitionID, BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                     "  	  BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                     "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                     "        ForecastMaturityDate) " +
                                     " Select '" + RequisitionID + "', BidBondInfoSeq, TenderMainSeq, Contractor, TaxId, Price, " +
                                     " 	      BidBondCategoryName, BidBondTypeName, CustodyTypeName, Amount, BillsNo, " +
                                     "        CheckBank, BillMaturityDate, CDExtensionTypenName, CDExtensionDeadlineDate, " +
                                     "        ForecastMaturityDate " +
                                     "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                                     "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq<>@BidBondInfoSeq ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);
                            
                            //履約、保固保證金附件
                            TmpSql = " Insert into FM7T_A67_F " +
                            		 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                            		 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                            		 " Select (Select isnull(max(AutoCounter),0) from FM7T_A67_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                            		 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                            		 "   from " + DBName + "dbo.V_BidBondInfo_Export " +
                            		 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);
                            
                            //歷史履約、保固保證金附件
                            TmpSql = " Insert into FM7T_A67_F " +
                            		 "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                            		 "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                            		 " Select (Select isnull(max(AutoCounter),0) from FM7T_A67_F)+ROW_NUMBER() over (order by TaxId), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                            		 " 	       AttFilePath, AttFileName, 0, 0, 1 " +
                            		 "   from " + DBName + "dbo.V_BidBondInfo_History_Export " +
                            		 "  Where TenderMainSeq=@TenderMainSeq and BidBondInfoSeq=@BidBondInfoSeq ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            ParameterList.Add(new SqlParameter("@BidBondInfoSeq", BidBondInfoSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);

                            //取得購案附件
                            TmpSql = " Insert into FM7T_A67_F " +
                                     "        (AutoCounter, RequisitionID, AccountID, MemberName, ProcessID, ProcessName, " +
                                     "  	  NFileName, OFileName, FileSize, DraftFlag, Type) " +
                                     " Select (Select isnull(max(AutoCounter),0) from FM7T_A67_F)+ROW_NUMBER() over (order by OrderNo), '" + RequisitionID + "', '" + ApplicantID + "', '" + ApplicantName + "', 'Start01', '填單', " +
                                     " 	       FullPath, FileName, 0, 0, 1 " +
                                     "   from " + DBName + "dbo.V_DocumentUpload_Export " +
                                     "  Where TenderMainSeq=@TenderMainSeq " +
                                     "  order by OrderNo ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@TenderMainSeq", TenderMainSeq.Replace("'", "")));
                            SqlCommand(connString, TmpSql, ParameterList);

                            //關聯表單
                            TmpSql = " exec dbo.SCSB_AssociationTable @Identify, @RequisitionID ";
                            ParameterList.Clear();
                            ParameterList.Add(new SqlParameter("@Identify", "A67"));
                            ParameterList.Add(new SqlParameter("@RequisitionID", RequisitionID));
                            SqlCommand(connString, TmpSql, ParameterList);
                        }
                    }
                    catch (Exception ex)
                    {
                        Response.Write(OutputJosn(5,"",""));
                        ErrorMsg = OutputJosn(99, ex.Message.ToString(), "");
                        return ErrorMsg;
                    }
                    break;
        }
        //起單跑流程
        using (SqlConnection InterfaceConnection = new SqlConnection(connString))
        {
            InterfaceConnection.Open();
            if (FM7Engine.Start(InterfaceConnection, RequisitionID, sDiagramGuid, ApplicantID, ApplicantDept, ref ErrorMsg) == FlowReturn.OK)
            {
                string TmpSql_99 = " exec [dbo].[SCSB_BreakLine] '','' ";
                ParameterList.Clear();
                SqlCommand(connString, TmpSql_99, ParameterList);




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

                //成功起單後判斷是否為重送表單
                if (OrgRequisitionID != "")
                {
                    TmpSql = " Insert into FSe7en_Tep_ResentList " +
                             "  	  (RequisitionID, OriginalRequisitionID) " +
                             " Select '" + RequisitionID + "', '" + OrgRequisitionID + "' ";
                    ParameterList.Clear();
                    SqlCommand(connString, TmpSql, ParameterList);

                }
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



	 public void LogRequest(){
	 
		  string connectionString =LoadCmdStr("\\\\Database\\\\Project\\\\SCSB\\\\Flow\\\\Connection\\\\BPM.xdbc.xmf",1);
		  string sCommand = "";
		  string returnData = "";
		  string requestUrl = Server.HtmlEncode(Request.RawUrl);
        using (SqlConnection  cnn = new SqlConnection(connectionString)){
				 cnn.Open();//連上資料庫
				 string strSQL = @"INSERT INTO [dbo].[SCSB_Log_RequestURL] ([RequestURL])  VALUES (@RequestURL)";
				 SqlCommand  myCommand = new SqlCommand (strSQL, cnn);
				 myCommand.Parameters.AddWithValue("@RequestURL", requestUrl);
				 myCommand.ExecuteNonQuery();
				 cnn.Close();
        }
		  return;
	 }

}