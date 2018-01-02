using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using SAP.Middleware.Connector;
using System.Text;
using System.Threading.Tasks;


public partial class _Default : System.Web.UI.Page
{
    public string PPlNo = "";
    public string Plant = "";
    public string LogMessage;
    public string Status;
    public string FileName;
    ClassSQL sql = new ClassSQL();
    WriteLog log = new WriteLog();
    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            Response.ContentType = "text/plain";
            PPlNo = Request.QueryString["PPL"];
            Plant = Request.QueryString["PLANT"];
            Status = Request.QueryString["STATUS"];
            // if (PPlNo != null && Plant != null)
            if (Plant != null)
            {
                if (Status == null)
                {
                    Status = "CREATE";
                }
                GetSAPData();
            }
            else
            {
                 Response.Write( "Plant Cannot be empty!!!");
                LogMessage = logMsg() + "Plant Cannot be empty!!!";
                FileName = getLogFilename();
                log.WriteLogFile(FileName, LogMessage , Plant);
            }
            Response.End();
        }
        catch(Exception ex)
        {
             Response.Write( ex.Message.ToString());
            LogMessage = logMsg() + ex.ToString();
            FileName = getLogFilename();
            log.WriteLogFile(FileName, LogMessage , Plant);
            return;
        }
    }
    public string getLogFilename()
    {
        string filename = "PPL_To_MES-" + Plant + " - " + DateTime.Now.ToString("yyyy-MM-dd hh_mm_ss")+".txt";
        return filename;
    }
    public string logMsg()
    {
        string log = "Program : Get PPL Details from SAP to MES - " + Plant + " - " + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") +" - Log Message : ";
        return log;
    }
    private static RfcDestination connectSAP()
    {
        RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("Salcomp");
        return SapRfcDestination;
    }
    public void GetSAPData()
    {
        try
        {
            if (Status == "CANCEL")
            {
                sql.ExecuteNonQuery("update [dbo].[PPLDetailsFromSAP] set Status='Canceled' where PPL='" + PPlNo + "'");
            }
            else if(Status == "CREATE")
            {
                RfcDestination SapRfcDestination = connectSAP();
                RfcSessionManager.BeginContext(SapRfcDestination);
                IRfcFunction funcPPL = SapRfcDestination.Repository.CreateFunction("ZMM_PPL_TO_MES");
               // funcPPL.SetValue("P_PPLNO", PPlNo);
                funcPPL.SetValue("P_WERKS", Plant);
                funcPPL.Invoke(SapRfcDestination);
                IRfcTable tblLog = funcPPL.GetTable("T_LOG");
                for (int j = 0; j < tblLog.Count; j++)
                {
                    if (tblLog[j].GetValue("MSGTY").ToString() == "S")
                    {
                        IRfcTable tblPPL = funcPPL.GetTable("T_PPL_INFO");
                        string PPL_no = " ";
                        for (int i = 0; i < tblPPL.Count; i++)
                        {
                            if (PPL_no != tblPPL[j].GetValue("PRELI").ToString())
                            {
                                string sqlCheck = "select count(*) from [dbo].[PPLDetailsFromSAP] where PPL='" + tblPPL[j].GetValue("PRELI").ToString() + "'";
                                var count = sql.ExecuteScalar(sqlCheck);
                                if (count != null)
                                {
                                    sql.ExecuteNonQuery("DELETE FROM [dbo].[PPLDetailsFromSAP] where PPL='" + tblPPL[j].GetValue("PRELI").ToString() + "'");
                                }
                            }
                            string sqlInsert = "INSERT INTO [dbo].[PPLDetailsFromSAP]" +
                                                "([PPL]" +
                                                ",[PplType]" +
                                                ",[Plant]" +
                                                ",[Product]" +
                                                ",[ProductionOrder]" +
                                                ",[OrderQty]" +
                                                ",[PoQtyUom]" +
                                                ",[MaterialGroup]" +
                                                ",[Material]" +
                                                ",[MaterialDesc]" +
                                                ",[Batch]" +
                                                ",[ReqQty]" +
                                                ",[ReqQtyUom]" +
                                                ",[Location]" +
                                                ",[ManufDate]" +
                                                ",[ExpDate]" +
                                                ",[SupplierCode]" +
                                                ",[SupplierName]" +
                                                ",[PplCreatedDate]" +
                                                ",[PplCreatedTime]" +
                                                ",[Status]" +
                                                ",[DeleteIndicator]" +
                                                ",[DocCreatedDate]" +
                                                ",[DocCreatedTime])" +
                                            "VALUES" +
                                                "('" + tblPPL[i].GetValue("PRELI").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("PPLTY").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("WERKS").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("MATNR").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("AUFNR").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("ORQTY").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("POUOM").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("EXTWG").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("IDNRK").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("MAKTX").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("CHARG").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("MENGE").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("REUOM").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("LGORT").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("HSDAT").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("VFDAT").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("LIFNR").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("NAME1").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("CRDAT").ToString() + "'" +
                                                ",'" + tblPPL[i].GetValue("CRTIM").ToString() + "'" +
                                                ",'Created'" +
                                                ",'0'" +
                                               ",'" + DateTime.Today.ToString("yyyy-MM-dd") + "'" +
                                                ",'" + DateTime.Now.ToString("hh:mm:ss tt") + "')";
                           
                            sql.ExecuteNonQuery(sqlInsert);
                            PPL_no = tblPPL[j].GetValue("PRELI").ToString();
                        }
                        IRfcTable tblAVL = funcPPL.GetTable("T_AVL_INFO");
                        for(int k=0; k < tblAVL.Count;  k++)
                        {
                            if(k == 0)
                            {
                                var check = sql.ExecuteScalar("select count(*) from [dbo].[ApprovedVendorListFromSAP] where product='" + tblAVL[k].GetValue("MATNR").ToString() + "'");
                                if(check != null)
                                {
                                    sql.ExecuteNonQuery("delete from [dbo].[ApprovedVendorListFromSAP] where product='" + tblAVL[k].GetValue("MATNR").ToString() + "'");
                                }
                            }
                            string SQLinsert = "INSERT INTO [dbo].[ApprovedVendorListFromSAP] " +
                                                 "([Plant]" +
                                                 ",[Product]" +
                                                 ",[Component]" +
                                                 ",[ApprovedVendor]" +
                                                 ",[Docdatetime])" +
                                             "VALUES" +
                                                 "('" + tblAVL[k].GetValue("WERKS").ToString() + "'" +
                                                 ",'" + tblAVL[k].GetValue("MATNR").ToString() + "'" +
                                                 ",'" + tblAVL[k].GetValue("IDNRK").ToString() + "'" +
                                                 ",'" + tblAVL[k].GetValue("LIFNR").ToString() + "'" +
                                                 ",'" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "')";
                            sql.ExecuteNonQuery(SQLinsert);
                        }
                        string log_m = tblPPL.Count.ToString() + " PPL Records and " + tblAVL.Count.ToString() + " AVL Records Successfully Exported!!!";
                         Response.Write( log_m);
                        LogMessage = logMsg() + log_m;
                        FileName = getLogFilename();
                        log.WriteLogFile(FileName, LogMessage,Plant);
                    }
                    else
                    {
                         Response.Write( tblLog[j].GetValue("MSGTX").ToString());
                        LogMessage = logMsg() + tblLog[j].GetValue("MSGTX").ToString();
                        FileName = getLogFilename();
                        log.WriteLogFile(FileName, LogMessage,Plant);
                    }

                }
                RfcSessionManager.EndContext(SapRfcDestination);
            }
        }
        catch(Exception ex)
        {
             Response.Write( ex.Message.ToString());
            LogMessage = logMsg() + ex.ToString();
            FileName = getLogFilename();
            log.WriteLogFile(FileName, LogMessage,Plant);
        }

    }

}