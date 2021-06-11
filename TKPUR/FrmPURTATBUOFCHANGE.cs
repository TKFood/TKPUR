using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using FastReport;
using FastReport.Data;
using System.Xml;

namespace TKPUR
{
    public partial class FrmPURTATBUOFCHANGE : Form
    {
        //測試ID = "170c1ac7-cdd2-46e1-bffa-aa37111ab514";
        //正式ID ="9c26cd05-861e-4e51-b090-d8e2fe3e685c"
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";
        string ID = "e45eda63-dc08-44a2-95af-99e95e7d0d9b";
        string DBNAME = "UOFTEST";


        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();


        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;
        public Report report1 { get; private set; }
        string REPORTID;
        string DELID;

        int ROWSINDEX = 0;
        int COLUMNSINDEX = 0;


        public FrmPURTATBUOFCHANGE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT [TA001] AS '請購單',[TA002] AS '請購單號',[TA006] AS '單頭備註'
                                    FROM [TK].dbo.PURTA
                                    WHERE TA001 LIKE '%{0}%' AND TA002 LIKE '%{1}%'
                                    ORDER BY  [TA001],[TA002]
                                    ",textBox1.Text.ToString().Trim(), textBox2.Text.ToString().Trim());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();

                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox7.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["請購單號"].Value.ToString().Trim();

                    textBox3.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["請購單號"].Value.ToString().Trim();


                }
                else
                {
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";

                }

                SearchPURTB(textBox7.Text, textBox8.Text);

                string MAXVERSIONS = GETMAXVERSIONSPURTATBCHAGE(textBox7.Text, textBox8.Text);
                textBox9.Text = (Convert.ToInt32(MAXVERSIONS) + 1).ToString();
            }
        }


        public void SearchPURTB(string TB001,string TB002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT 
                                    [TB003] AS '請購序號'
                                    ,[TB004] AS '品號'
                                    ,[TB005] AS '品名'
                                    ,[TB009] AS '請購教量'
                                    ,[TB011] AS '需求日'
                                    ,[TB012] AS '單身備註'
                                    FROM [TK].dbo.PURTB
                                    WHERE TB001='{0}' AND TB002='{1}'
                                    ORDER BY [TB003]
                                    ",TB001,TB002 );

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["ds"];
                        dataGridView2.AutoResizeColumns();

                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public string GETMAXVERSIONSPURTATBCHAGE(string TA001,string TA002)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                SqlTransaction tran;
                SqlCommand cmd = new SqlCommand();
                DataSet ds = new DataSet();
                string VERSIONS;

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;

                sqlConn = new SqlConnection(connectionString);
                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();

               
                sbSql.AppendFormat(@"  
                                    SELECT
                                    MAX([VERSIONS]) AS VERSIONS,[TA001],[TA002]
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001]='{0}' AND [TA002]='{1}'
                                    GROUP BY [TA001],[TA002]
                                    ", TA001,TA002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    VERSIONS = ds.Tables["ds"].Rows[0]["VERSIONS"].ToString();
                    return VERSIONS;

                }
                else
                {
                    return "0";
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDPURTATBCHAGE(string TA001,string TA002,string VERSIONS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                                
                sbSql.AppendFormat(@"  
                                    INSERT INTO [TKPUR].[dbo].[PURTATBCHAGE]
                                    ([VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002])

                                    SELECT '{2}',[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB007],[TB009],[TB010],[TB011],[TB012]
                                    ,USER_GUID
                                    ,[NAME]
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,MA002
                                    FROM 
                                    (
                                    SELECT [TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB007],[TB009],[TB010],[TB011],[TB012]
                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TA012) AS 'MV002'
                                    ,(SELECT TOP 1 MA002 FROM [TK].dbo.PURMA WHERE MA001=TB010) AS 'MA002'
                                    FROM [TK].dbo.PURTB,[TK].dbo.PURTA
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TA012 COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA001='{0}' AND TA002='{1}'
                                    ) AS TEMP
                                    ", TA001,  TA002,  VERSIONS);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SEARCHPURTATBCHAGEVERSIONS(string TA001,string TA002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [VERSIONS] AS '變更', [TA001] AS '請購單',[TA002] AS '請購單號'
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001] LIKE '%{0}%' AND [TA002] LIKE '%{1}%'
                                    GROUP BY [VERSIONS],[TA001],[TA002]
                                    ORDER BY [TA001],[TA002],[VERSIONS] DESC
                                    ", TA001, TA002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds.Tables["ds"];
                        dataGridView3.AutoResizeColumns();

                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView4.DataSource = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox10.Text = row.Cells["變更"].Value.ToString().Trim();
                    textBox11.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox12.Text = row.Cells["請購單號"].Value.ToString().Trim();



                }
                else
                {
                    textBox7.Text = "";
                    textBox8.Text = "";
                    

                }
                SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);


            }
        }

        public void SEARCHPURTATBCHAGE(string TA001,string TA002,string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [VERSIONS] AS '變更版號'
                                    ,[TA001] AS '請購單'
                                    ,[TA002] AS '請購單號'
                                    ,[TA006] AS '單頭備註'
                                    ,[TB003] AS '請購序號'
                                    ,[TB004] AS '品號'
                                    ,[TB005] AS '品名'
                                    ,[TB009] AS '請購教量'
                                    ,[TB011] AS '需求日'
                                    ,[TB012] AS '單身備註'
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}'
                                    ORDER BY [TA001] ,[TA002],[TB003]
                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds.Tables["ds"];
                        dataGridView4.AutoResizeColumns();

                    }

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBox13.Text = row.Cells["單頭備註"].Value.ToString().Trim();
                    textBox14.Text = row.Cells["請購序號"].Value.ToString().Trim();
                    textBox15.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox16.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox17.Text = row.Cells["請購教量"].Value.ToString().Trim();
                    textBox18.Text = row.Cells["需求日"].Value.ToString().Trim();
                    textBox19.Text = row.Cells["單身備註"].Value.ToString().Trim();

                }
                else
                {
                    textBox13.Text = "";
                    textBox14.Text = "";
                    textBox15.Text = "";
                    textBox16.Text = "";
                    textBox17.Text = "";
                    textBox18.Text = "";
                    textBox19.Text = "";

                }
               

            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            textBox16.Text = SEARCHINVMB(textBox15.Text.Trim());
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            textBox22.Text = SEARCHINVMB(textBox21.Text.Trim());
        }

        public string SEARCHINVMB(string MB001)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                SqlTransaction tran;
                SqlCommand cmd = new SqlCommand();
                DataSet ds = new DataSet();
                string MB002;

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;

                sqlConn = new SqlConnection(connectionString);
                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT MB002 
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001='{0}'
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    MB002 = ds.Tables["ds"].Rows[0]["MB002"].ToString();
                    return MB002;

                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEPURTATBCHAGETA006(string TA001, string TA002, string VERSIONS,string TA006)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    UPDATE  [TKPUR].[dbo].[PURTATBCHAGE]
                                    SET [TA006]='{3}'
                                    WHERE  [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}'
                                    ", TA001, TA002, VERSIONS,TA006);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }


        public void UPDATEPURTATBCHAGE(string TA001, string TA002, string VERSIONS, string TB003, string TB004, string TB005, string TB009, string TB011, string TB012)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    UPDATE  [TKPUR].[dbo].[PURTATBCHAGE]
                                    SET [TB004]='{4}',[TB005]='{5}',[TB009]='{6}',[TB011]='{7}',[TB012]='{8}'
                                    WHERE  [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}' AND TB003='{3}'
                                   
                                    ", TA001, TA002, VERSIONS, TB003,  TB004,  TB005,  TB009,  TB011,  TB012);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDPURTATBCHAGEDETAIL(string VERSIONS, string TA001, string TA002, string TA006, string TB003, string TB004, string TB005, string TB009, string TB011, string TB012)
        {
            DataSet dsUSER = new DataSet();
            DataSet dsINVMB = new DataSet();

            dsUSER = SEARCHUSER(TA001, TA002, VERSIONS);
            dsINVMB = SEARCHINVMBALL(TB004);

            string TA003 = dsUSER.Tables[0].Rows[0]["TA003"].ToString();
            string TA012 = dsUSER.Tables[0].Rows[0]["TA012"].ToString();
            string USER_GUID = dsUSER.Tables[0].Rows[0]["USER_GUID"].ToString();
            string NAME = dsUSER.Tables[0].Rows[0]["NAME"].ToString();
            string GROUP_ID = dsUSER.Tables[0].Rows[0]["GROUP_ID"].ToString();
            string TITLE_ID = dsUSER.Tables[0].Rows[0]["TITLE_ID"].ToString();

            string TB007 = dsINVMB.Tables[0].Rows[0]["MB004"].ToString();
            string TB010 = dsINVMB.Tables[0].Rows[0]["MB032"].ToString();
            string MA002 = dsINVMB.Tables[0].Rows[0]["MA002"].ToString();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                     INSERT INTO [TKPUR].[dbo].[PURTATBCHAGE]
                                    ([VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}')
                                    ", VERSIONS,TA001,TA002, TA003, TA006,TA012,TB003,TB004,TB005,TB007,TB009,TB010,TB011,TB012,USER_GUID,NAME,GROUP_ID,TITLE_ID,MA002);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public DataSet SEARCHUSER(string TA001,string TA002,string VERSIONS)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                SqlTransaction tran;
                SqlCommand cmd = new SqlCommand();
                DataSet ds = new DataSet();
                string MB002;

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;

                sqlConn = new SqlConnection(connectionString);
                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TOP 1
                                    [VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002]
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE ISNULL(TA012,'')<>''
                                    AND TA001='{0}' AND TA002='{1}' AND [VERSIONS]='{2}'
                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                   
                    return ds;

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataSet SEARCHINVMBALL(string MB001)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                SqlTransaction tran;
                SqlCommand cmd = new SqlCommand();
                DataSet ds = new DataSet();
                string MB002;

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;

                sqlConn = new SqlConnection(connectionString);
                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT TOP 1 MB001,MB002,MB003,MB004,MB032,MA002
                                    FROM [TK].dbo.INVMB
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=MB032
                                    WHERE MB001='{0}'
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {

                    return ds;

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }


        public void ADDTB_WKF_EXTERNAL_TASK(string TA001, string TA002,string VERSIONS)
        {
            DataTable DT = SEARCHPURTAPURTB(TA001, TA002, VERSIONS);
            DataTable DTUPFDEP = SEARCHUOFDEP(DT.Rows[0]["TA012"].ToString());

            string EXTERNAL_FORM_NBR= DT.Rows[0]["TA001"].ToString().Trim()+ DT.Rows[0]["TA002"].ToString().Trim() + DT.Rows[0]["VERSIONS"].ToString().Trim();

            string account = DT.Rows[0]["TA012"].ToString();
            string groupId = DT.Rows[0]["GROUP_ID"].ToString();
            string jobTitleId = DT.Rows[0]["TITLE_ID"].ToString();
            string fillerName = DT.Rows[0]["NAME"].ToString();
            string fillerUserGuid = DT.Rows[0]["USER_GUID"].ToString();


            string DEPNAME = DTUPFDEP.Rows[0]["DEPNAME"].ToString();
            string DEPNO= DTUPFDEP.Rows[0]["DEPNO"].ToString();

            int rowscounts = 0;

            XmlDocument xmlDoc = new XmlDocument();
            //建立根節點
            XmlElement Form = xmlDoc.CreateElement("Form");

            //正式的id
            Form.SetAttribute("formVersionId", ID);

            Form.SetAttribute("urgentLevel", "2");
            //加入節點底下
            xmlDoc.AppendChild(Form);

            ////建立節點Applicant
            XmlElement Applicant = xmlDoc.CreateElement("Applicant");
            Applicant.SetAttribute("account", account);
            Applicant.SetAttribute("groupId", groupId);
            Applicant.SetAttribute("jobTitleId", jobTitleId);
            //加入節點底下
            Form.AppendChild(Applicant);

            //建立節點 Comment
            XmlElement Comment = xmlDoc.CreateElement("Comment");
            Comment.InnerText = "申請者意見";
            //加入至節點底下
            Applicant.AppendChild(Comment);

            //建立節點 FormFieldValue
            XmlElement FormFieldValue = xmlDoc.CreateElement("FormFieldValue");
            //加入至節點底下
            Form.AppendChild(FormFieldValue);

            //建立節點FieldItem
            //ID 表單編號	
            XmlElement FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "ID");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //DEPNO 變更版本	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "DEPNO");
            FieldItem.SetAttribute("fieldValue", DEPNAME);
            FieldItem.SetAttribute("realValue", DEPNO);
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);


            //建立節點FieldItem
            //VERSIONS 變更版本	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "VERSIONS");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["VERSIONS"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA001 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA001");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA001"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA002 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA002"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA003 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA003");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA003"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA012 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA012");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA012"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //MV002 姓名	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "MV002");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["NAME"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TA006 單頭備註	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TA006");
            FieldItem.SetAttribute("fieldValue", DT.Rows[0]["TA006"].ToString());
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點FieldItem
            //TB 表單編號	
            FieldItem = xmlDoc.CreateElement("FieldItem");
            FieldItem.SetAttribute("fieldId", "TB");
            FieldItem.SetAttribute("fieldValue", "");
            FieldItem.SetAttribute("realValue", "");
            FieldItem.SetAttribute("enableSearch", "True");
            FieldItem.SetAttribute("fillerName", fillerName);
            FieldItem.SetAttribute("fillerUserGuid", fillerUserGuid);
            FieldItem.SetAttribute("fillerAccount", account);
            FieldItem.SetAttribute("fillSiteId", "");
            //加入至members節點底下
            FormFieldValue.AppendChild(FieldItem);

            //建立節點 DataGrid
            XmlElement DataGrid = xmlDoc.CreateElement("DataGrid");
            //DataGrid 加入至 TB 節點底下
            XmlNode TB = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='TB']");
            TB.AppendChild(DataGrid);


            foreach (DataRow od in DT.Rows)
            {
                // 新增 Row
                XmlElement Row = xmlDoc.CreateElement("Row");
                Row.SetAttribute("order", (rowscounts).ToString());

                //Row	TB004
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB004");
                Cell.SetAttribute("fieldValue", od["TB004"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB005
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB005");
                Cell.SetAttribute("fieldValue", od["TB005"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                ////Row	TB006
                //Cell = xmlDoc.CreateElement("Cell");
                //Cell.SetAttribute("fieldId", "TB006");
                //Cell.SetAttribute("fieldValue", od["TB006"].ToString());
                //Cell.SetAttribute("realValue", "");
                //Cell.SetAttribute("customValue", "");
                //Cell.SetAttribute("enableSearch", "True");
                ////Row
                //Row.AppendChild(Cell);

                //Row	TB007
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB007");
                Cell.SetAttribute("fieldValue", od["TB007"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB009
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB009");
                Cell.SetAttribute("fieldValue", od["TB009"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB011
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB011");
                Cell.SetAttribute("fieldValue", od["TB011"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB010
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB010");
                Cell.SetAttribute("fieldValue", od["TB010"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	MA002
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "MA002");
                Cell.SetAttribute("fieldValue", od["MA002"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB012
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB012");
                Cell.SetAttribute("fieldValue", od["TB012"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                rowscounts = rowscounts + 1;

                XmlNode DataGridS = xmlDoc.SelectSingleNode("./Form/FormFieldValue/FieldItem[@fieldId='TB']/DataGrid");
                DataGridS.AppendChild(Row);

            }

            ////用ADDTACK，直接啟動起單
            //ADDTACK(Form);

            //ADD TO DB
            string connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ToString();

            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                }
            }
            catch
            {

            }
            finally
            {

            }





        }

        public DataTable SEARCHPURTAPURTB(string TA001, string TA002,string VERSIONS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    [VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002]
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}'
                                    ORDER BY [VERSIONS],[TA001],[TA002],[TB003]
                              
                                    ", TA001, TA002, VERSIONS);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataTable SEARCHUOFDEP(string ACCOUNT)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [GROUP_NAME] AS 'DEPNAME'
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]+','+[GROUP_NAME]+',False' AS 'DEPNO'
                                    ,[TB_EB_USER].[USER_GUID]
                                    ,[ACCOUNT]
                                    ,[NAME]
                                    ,[TB_EB_EMPL_DEP].[GROUP_ID]
                                    ,[TITLE_ID]     
                                    ,[GROUP_NAME]
                                    ,[GROUP_CODE]
                                    FROM [192.168.1.223].[UOF].[dbo].[TB_EB_USER],[192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP],[192.168.1.223].[UOF].[dbo].[TB_EB_GROUP]
                                    WHERE [TB_EB_USER].[USER_GUID]=[TB_EB_EMPL_DEP].[USER_GUID]
                                    AND [TB_EB_EMPL_DEP].[GROUP_ID]=[TB_EB_GROUP].[GROUP_ID]
                                    AND ISNULL([TB_EB_GROUP].[GROUP_CODE],'')<>''
                                    AND [ACCOUNT]='{0}'
                              
                                    ", ACCOUNT);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void SETNULL()
        {
            textBox20.Text = null;
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }


        private void button4_Click(object sender, EventArgs e)
        {
            ADDPURTATBCHAGE(textBox7.Text, textBox8.Text, textBox9.Text);

            textBox9.Text = (Convert.ToInt32(textBox9.Text) + 1).ToString();

            MessageBox.Show("轉入完成");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHPURTATBCHAGEVERSIONS(textBox3.Text.Trim(), textBox4.Text.Trim());
        }



        private void button5_Click(object sender, EventArgs e)
        {
            UPDATEPURTATBCHAGETA006(textBox11.Text, textBox12.Text, textBox10.Text, textBox13.Text);

            SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);

            MessageBox.Show("完成");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UPDATEPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text, textBox14.Text, textBox15.Text, textBox16.Text, textBox17.Text, textBox18.Text, textBox19.Text);

            SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);

            MessageBox.Show("完成");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ADDPURTATBCHAGEDETAIL(textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox20.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text);
                                     
            SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);

            SETNULL();
            MessageBox.Show("完成");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            ADDTB_WKF_EXTERNAL_TASK(textBox11.Text.Trim(), textBox12.Text.Trim(),textBox10.Text.Trim());
        }
        #endregion


    }
}
