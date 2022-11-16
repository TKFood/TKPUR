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
using TKITDLL;

namespace TKPUR
{
    public partial class FrmPURTATBUOFCHANGE : Form
    {
        //測試ID = "";
        //正式ID ="c8441ee2-d3bb-4c30-b731-f19a7916566f"
        //測試DB DBNAME = "UOFTEST";
        //正式DB DBNAME = "UOF";
        string DBNAME = "UOF";



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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                                
                sbSql.AppendFormat(@"  
                                    INSERT INTO [TKPUR].[dbo].[PURTATBCHAGE]
                                    ([VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB006],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002])

                                    SELECT '{3}',[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB006],[TB007],[TB009],[TB010],[TB011],[TB012]
                                    ,USER_GUID
                                    ,[NAME]
                                    ,(SELECT TOP 1 GROUP_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'GROUP_ID'
                                    ,(SELECT TOP 1 TITLE_ID FROM [192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP] WHERE [TB_EB_EMPL_DEP].USER_GUID=TEMP.USER_GUID) AS 'TITLE_ID'
                                    ,MA002
                                    FROM 
                                    (
                                    SELECT [TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB006],[TB007],[TB009],[TB010],[TB011],[TB012]
                                    ,[TB_EB_USER].USER_GUID,NAME
                                    ,(SELECT TOP 1 MV002 FROM [TK].dbo.CMSMV WHERE MV001=TA012) AS 'MV002'
                                    ,(SELECT TOP 1 MA002 FROM [TK].dbo.PURMA WHERE MA001=TB010) AS 'MA002'
                                    FROM [TK].dbo.PURTB,[TK].dbo.PURTA
                                    LEFT JOIN [192.168.1.223].[{0}].[dbo].[TB_EB_USER] ON [TB_EB_USER].ACCOUNT= TA012 COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA001='{1}' AND TA002='{2}'
                                    ) AS TEMP
                                    ", DBNAME, TA001,  TA002,  VERSIONS);

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
            dataGridView4.DataSource = null;
            dataGridView4.DataSource = null;
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
                    textBox10.Text = "";
                    textBox11.Text = "";
                    textBox12.Text = "";


                }

                SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);

                textBox26.Text= SERACHlDOC_NBR(textBox11.Text+ textBox12.Text+textBox10.Text);
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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
            DataSet dsINVMB = new DataSet();
            
            dsINVMB = SEARCHINVMBALL(TB004);


            string TB006 = dsINVMB.Tables[0].Rows[0]["MB003"].ToString();
            string TB007 = dsINVMB.Tables[0].Rows[0]["MB004"].ToString();
            string TB010 = dsINVMB.Tables[0].Rows[0]["MB032"].ToString();
            string MA002 = dsINVMB.Tables[0].Rows[0]["MA002"].ToString();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    UPDATE  [TKPUR].[dbo].[PURTATBCHAGE]
                                    SET [TB004]='{4}',[TB005]='{5}',[TB006]='{6}',[TB009]='{7}',[TB011]='{8}',[TB012]='{9}'
                                    WHERE  [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}' AND TB003='{3}'
                                   
                                    ", TA001, TA002, VERSIONS, TB003,  TB004,  TB005, TB006,  TB009,  TB011,  TB012);

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

            string TB006 = dsINVMB.Tables[0].Rows[0]["MB003"].ToString();
            string TB007 = dsINVMB.Tables[0].Rows[0]["MB004"].ToString();
            string TB010 = dsINVMB.Tables[0].Rows[0]["MB032"].ToString();
            string MA002 = dsINVMB.Tables[0].Rows[0]["MA002"].ToString();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                     INSERT INTO [TKPUR].[dbo].[PURTATBCHAGE]
                                    ([VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB006],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}')
                                    ", VERSIONS,TA001,TA002, TA003, TA006,TA012,TB003,TB004,TB005, TB006, TB007,TB009,TB010,TB011,TB012,USER_GUID,NAME,GROUP_ID,TITLE_ID,MA002);

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


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

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


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
            string PURCHID = SEARCHFORM_UOF_VERSION_ID("PUR20.請購單變更單");
            //string PURCHID = SEARCHFORM_VERSION_ID("PUR20.請購單變更單");


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
            Form.SetAttribute("formVersionId", PURCHID);

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

                //Row	TB003
                XmlElement Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB003");
                Cell.SetAttribute("fieldValue", od["TB003"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

                //Row	TB004
                Cell = xmlDoc.CreateElement("Cell");
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

                //Row	TB006
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TB006");
                Cell.SetAttribute("fieldValue", od["TB006"].ToString());
                Cell.SetAttribute("realValue", "");
                Cell.SetAttribute("customValue", "");
                Cell.SetAttribute("enableSearch", "True");
                //Row
                Row.AppendChild(Cell);

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
             
                //TD001002003 	
                Cell = xmlDoc.CreateElement("Cell");
                Cell.SetAttribute("fieldId", "TD001002003");
                Cell.SetAttribute("fieldValue", od["TD001002003"].ToString());
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

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);
            


            StringBuilder queryString = new StringBuilder();




            queryString.AppendFormat(@" INSERT INTO [{0}].dbo.TB_WKF_EXTERNAL_TASK
                                         (EXTERNAL_TASK_ID,FORM_INFO,STATUS,EXTERNAL_FORM_NBR)
                                        VALUES (NEWID(),@XML,2,'{1}')
                                        ", DBNAME, EXTERNAL_FORM_NBR);

            try
            {
                using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
                {

                    SqlCommand command = new SqlCommand(queryString.ToString(), connection);
                    command.Parameters.Add("@XML", SqlDbType.NVarChar).Value = Form.OuterXml;

                    command.Connection.Open();

                    int count = command.ExecuteNonQuery();

                    connection.Close();
                    connection.Dispose();

                    MessageBox.Show("已送出UOF");

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    [VERSIONS],[TA001],[TA002],[TA003],[TA006],[TA012],[TB003],[TB004],[TB005],[TB006],[TB007],[TB009],[TB010],[TB011],[TB012],[USER_GUID],[NAME],[GROUP_ID],[TITLE_ID],[MA002]
                                    ,(SELECT TD001+' '+TD002+' '+TD003+CHAR(10) FROM [TK].dbo.PURTD WHERE  TD026=[TA001] AND TD027=[TA002] AND TD028=[TB003] FOR XML PATH('')) AS TD001002003

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
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
                                    FROM [192.168.1.223].[{0}].[dbo].[TB_EB_USER],[192.168.1.223].[{0}].[dbo].[TB_EB_EMPL_DEP],[192.168.1.223].[{0}].[dbo].[TB_EB_GROUP]
                                    WHERE [TB_EB_USER].[USER_GUID]=[TB_EB_EMPL_DEP].[USER_GUID]
                                    AND [TB_EB_EMPL_DEP].[GROUP_ID]=[TB_EB_GROUP].[GROUP_ID]
                                    AND ISNULL([TB_EB_GROUP].[GROUP_CODE],'')<>''
                                    AND [ACCOUNT]='{1}'
                              
                                    ", DBNAME, ACCOUNT);


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

        public string SERACHlDOC_NBR(string  EXTERNAL_FORM_NBR)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT TOP 1 EXTERNAL_FORM_NBR,DOC_NBR
                                    FROM [{0}].[dbo].[TB_WKF_EXTERNAL_TASK]
                                    WHERE EXTERNAL_FORM_NBR='{1}'
                                    ORDER BY DOC_NBR DESC
                                    ", DBNAME, EXTERNAL_FORM_NBR);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["DOC_NBR"].ToString();

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

        public string SERACHlDOC_NBR_CHECK(string EXTERNAL_FORM_NBR)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT TOP 1 EXTERNAL_FORM_NBR,[TB_WKF_EXTERNAL_TASK].DOC_NBR,TB_WKF_TASK.TASK_RESULT
                                    FROM [UOF].[dbo].[TB_WKF_EXTERNAL_TASK],[UOF].[dbo].TB_WKF_TASK 
                                    WHERE [TB_WKF_EXTERNAL_TASK].DOC_NBR=TB_WKF_TASK.DOC_NBR
                                    AND ISNULL(TB_WKF_TASK.TASK_RESULT,'-1') NOT IN ('0','1','2')
                                    AND EXTERNAL_FORM_NBR LIKE '%{0}%'
                                    ORDER BY DOC_NBR DESC
                                    ", EXTERNAL_FORM_NBR);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return "Y";

                }
                else
                {
                    return "N";
                }

            }
            catch
            {
                return "N";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void DELPURTATBCHAGE(string VERSIONS, string TA001, string TA002, string TB003)
        {
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                     DELETE [TKPUR].[dbo].[PURTATBCHAGE]
                                     WHERE [VERSIONS]='{0}' AND [TA001]='{1}' AND [TA002]='{2}' AND [TB003]='{3}'
                                    ", VERSIONS, TA001, TA002,  TB003);

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

        public void SETNULL()
        {
            textBox20.Text = null;
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;
        }

        public string SEARCHFORM_UOF_VERSION_ID(string FORM_NAME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                   SELECT TOP 1 RTRIM(LTRIM(TB_WKF_FORM_VERSION.FORM_VERSION_ID)) FORM_VERSION_ID,TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION,TB_WKF_FORM_VERSION.ISSUE_CTL
                                    ,TB_WKF_FORM.FORM_NAME
                                    FROM [UOF].dbo.TB_WKF_FORM_VERSION,[UOF].dbo.TB_WKF_FORM
                                    WHERE 1=1
                                    AND TB_WKF_FORM_VERSION.FORM_ID=TB_WKF_FORM.FORM_ID
                                    AND TB_WKF_FORM_VERSION.ISSUE_CTL=1
                                    AND FORM_NAME='{0}'
                                    ORDER BY TB_WKF_FORM_VERSION.FORM_ID,TB_WKF_FORM_VERSION.VERSION DESC

                                    ", FORM_NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["FORM_VERSION_ID"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }
        public string SEARCHFORM_VERSION_ID(string FORM_NAME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                //sqlConn = new SqlConnection(connectionString);

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT 
                                    RTRIM(LTRIM([FORM_VERSION_ID])) AS FORM_VERSION_ID
                                    ,[FORM_NAME]
                                    FROM [TKIT].[dbo].[UOF_FORM_VERSION_ID]
                                    WHERE [FORM_NAME]='{0}'
                                    ", FORM_NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");


                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"].Rows[0]["FORM_VERSION_ID"].ToString();
                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return "";
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void SEARCHPURTATBCHAGEDETAILS(string TA001, string TA002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



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
                                    WHERE [TA001] LIKE '%{0}%' AND [TA002] LIKE '%{1}%'
                                    ORDER BY [VERSIONS] DESC,[TA001] ,[TA002],[TB003]
                                    ", TA001, TA002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView5.DataSource = ds.Tables["ds"];
                        dataGridView5.AutoResizeColumns();

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

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {          

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    textBox27.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox28.Text = row.Cells["請購單號"].Value.ToString().Trim();
                    textBox29.Text = row.Cells["變更版號"].Value.ToString().Trim();
                    
                }
                else
                {
                    textBox27.Text = "";
                    textBox28.Text = "";
                    textBox29.Text = "";
                   

                }

                SEARCHPURTE(textBox27.Text, textBox28.Text, textBox29.Text);

               
            }
        }


        public void NEWPURTEPURTF(string TA001,string TA002,string VERSIONS)
        {

            //A311 20221101011 1
            //檢查請購變更單的採購單，是否有採購變更單未核準
            DataTable DTCHECKPURTEPURTF = CHECKPURTEPURTF(TA001, TA002, VERSIONS);

            if(DTCHECKPURTEPURTF== null)
            {
                //找出請購變更單有幾張採購單，要1對多
                DataTable DTPURTCPURTD = SEARCHPURTCPURTD(TA001, TA002, VERSIONS);
                //DataTable DTPURTCPURTD = SEARCHPURTCPURTD("A312", "20221116001", "2");
                DataTable DTOURTE = new DataTable();

                //找出採購單跟最大的版次
                if (DTPURTCPURTD.Rows.Count > 0)
                {
                    DTOURTE = FINDPURTE(DTPURTCPURTD);
                }

                //新增採購變更單
                if (DTOURTE.Rows.Count > 0)
                {
                    ADDTOPURTEPURTF(DTOURTE);
                }
            }
            else
            {
                StringBuilder MESS = new StringBuilder();
                foreach(DataRow DR in DTCHECKPURTEPURTF.Rows)
                {
                    MESS.AppendFormat(@" 採購變更單:"+ DR["TE001"].ToString()+" "+ DR["TE002"].ToString() + ""+"變更版次:" + DR["TE003"].ToString()+" 沒有核準 ");
                }

                MessageBox.Show(MESS.ToString());
            }
           

        }

        public DataTable CHECKPURTEPURTF(string TA001, string TA002, string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TE001,TE002,TE003
                                    FROM [TK].dbo.PURTE
                                    WHERE TE017 IN ('N')
                                    AND TE001+TE002 IN 
                                    (
                                    SELECT TD001+TD002
                                    FROM [TK].dbo.PURTD
                                    WHERE TD026+TD027+TD028 IN 
                                    (
                                    SELECT TA001+TA002+TB003
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE  TA001='{0}' AND TA002='{1}' AND VERSIONS='{2}'
                                    )
                                    GROUP BY  TD001,TD002
                                    )

                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    return ds.Tables["TEMPds1"];
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

            }



        }

        public DataTable SEARCHPURTCPURTD(string TA001, string TA002, string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                  
                                    SELECT TD001,TD002,'{0}' TA001,'{1}' TA002,'{2}' VERSIONS
                                    FROM [TK].dbo.PURTD
                                    WHERE TD026+TD027+TD028 IN 
                                    (
                                    SELECT TA001+TA002+TB003
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE  TA001='{0}' AND TA002='{1}' AND VERSIONS='{2}'
                                    )
                                    GROUP BY  TD001,TD002

                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {

                    return ds.Tables["TEMPds1"];
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

            }

           

        }

        public DataTable FINDPURTE(DataTable DTTEMP)
        {
            DataTable DT = new DataTable();
            DT.Clear();
            DT.Columns.Add("TE001");
            DT.Columns.Add("TE002");
            DT.Columns.Add("TE003");
            DT.Columns.Add("TA001");
            DT.Columns.Add("TA002");
            DT.Columns.Add("VERSIONS");



            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            string TE001 = null;
            string TE002 = null;
            string TA001 = null;
            string TA002 = null;
            string VERSIONS = null;

            if (DTTEMP.Rows.Count>0)
            {
                foreach(DataRow DR in DTTEMP.Rows)
                {

                    TE001 = DR["TD001"].ToString();
                    TE002 = DR["TD002"].ToString();
                    TA001 = DR["TA001"].ToString();
                    TA002 = DR["TA002"].ToString();
                    VERSIONS = DR["VERSIONS"].ToString();

                    try
                    {
                        //20210902密
                        Class1 TKID = new Class1();//用new 建立類別實體
                        SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                        //資料庫使用者密碼解密
                        sqlsb.Password = TKID.Decryption(sqlsb.Password);
                        sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                        String connectionString;
                        sqlConn = new SqlConnection(sqlsb.ConnectionString);



                        sbSql.Clear();
                        sbSqlQuery.Clear();


                        sbSql.AppendFormat(@"  
                                            SELECT TOP 1 TE001,TE002,TE003
                                            FROM [TK].dbo.PURTE
                                            WHERE TE001='{0}' AND TE002='{1}'
                                            ORDER BY TE001 DESC,TE002 DESC,TE003 DESC

                                             ", TE001, TE002);

                        adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                        sqlCmdBuilder = new SqlCommandBuilder(adapter);
                        sqlConn.Open();
                        ds.Clear();
                        adapter.Fill(ds, "TEMPds1");
                        sqlConn.Close();


                        if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                        {
                            int serno = Convert.ToInt16(ds.Tables["TEMPds1"].Rows[0]["TE003"].ToString());
                            serno = serno + 1;
                            string temp = serno.ToString();
                            temp = temp.PadLeft(4, '0');

                            DataRow NEWDR = DT.NewRow();
                            NEWDR["TE001"] = TE001;
                            NEWDR["TE002"] = TE002;
                            NEWDR["TE003"] = temp;
                            NEWDR["TA001"] = TA001;
                            NEWDR["TA002"] = TA002;
                            NEWDR["VERSIONS"] = VERSIONS;
                            DT.Rows.Add(NEWDR);

                        }
                        else
                        {
                            DataRow NEWDR = DT.NewRow();
                            NEWDR["TE001"] = TE001;
                            NEWDR["TE002"] = TE002;
                            NEWDR["TE003"] = "0001";
                            NEWDR["TA001"] = TA001;
                            NEWDR["TA002"] = TA002;
                            NEWDR["VERSIONS"] = VERSIONS;
                            DT.Rows.Add(NEWDR);

                        }


                    }
                    catch
                    {
                        return null;
                    }
                    finally
                    {

                    }
                }

                return DT;
            }
            else
            {
                return null;
            }
        }


        public void ADDTOPURTEPURTF(DataTable NEWPURTEPURTF)
        {
            if(NEWPURTEPURTF.Rows.Count>0)
            {
                try
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);



                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    foreach(DataRow DR in NEWPURTEPURTF.Rows)
                    {
                        sbSql.AppendFormat(@"  
                                   
                                            INSERT INTO [TK].[dbo].[PURTF]
                                            (
                                            [COMPANY]
                                            ,[CREATOR]
                                            ,[USR_GROUP]
                                            ,[CREATE_DATE]
                                            ,[MODIFIER]
                                            ,[MODI_DATE]
                                            ,[FLAG]
                                            ,[CREATE_TIME]
                                            ,[MODI_TIME]
                                            ,[TRANS_TYPE]
                                            ,[TRANS_NAME]
                                            ,[sync_date]
                                            ,[sync_time]
                                            ,[sync_mark]
                                            ,[sync_count]
                                            ,[DataUser]
                                            ,[DataGroup]
                                            ,[TF001]
                                            ,[TF002]
                                            ,[TF003]
                                            ,[TF004]
                                            ,[TF005]
                                            ,[TF006]
                                            ,[TF007]
                                            ,[TF008]
                                            ,[TF009]
                                            ,[TF010]
                                            ,[TF011]
                                            ,[TF012]
                                            ,[TF013]
                                            ,[TF014]
                                            ,[TF015]
                                            ,[TF016]
                                            ,[TF017]
                                            ,[TF018]
                                            ,[TF019]
                                            ,[TF020]
                                            ,[TF021]
                                            ,[TF022]
                                            ,[TF023]
                                            ,[TF024]
                                            ,[TF025]
                                            ,[TF026]
                                            ,[TF027]
                                            ,[TF028]
                                            ,[TF029]
                                            ,[TF030]
                                            ,[TF031]
                                            ,[TF032]
                                            ,[TF033]
                                            ,[TF034]
                                            ,[TF035]
                                            ,[TF036]
                                            ,[TF037]
                                            ,[TF038]
                                            ,[TF039]
                                            ,[TF040]
                                            ,[TF041]
                                            ,[TF104]
                                            ,[TF105]
                                            ,[TF106]
                                            ,[TF107]
                                            ,[TF108]
                                            ,[TF109]
                                            ,[TF110]
                                            ,[TF111]
                                            ,[TF112]
                                            ,[TF113]
                                            ,[TF114]
                                            ,[TF118]
                                            ,[TF119]
                                            ,[TF120]
                                            ,[TF121]
                                            ,[TF122]
                                            ,[TF123]
                                            ,[TF124]
                                            ,[TF125]
                                            ,[TF126]
                                            ,[TF127]
                                            ,[TF128]
                                            ,[TF129]
                                            ,[TF130]
                                            ,[TF131]
                                            ,[TF132]
                                            ,[TF133]
                                            ,[TF134]
                                            ,[TF135]
                                            ,[TF136]
                                            ,[TF137]
                                            ,[TF138]
                                            ,[TF139]
                                            ,[TF140]
                                            ,[TF141]
                                            ,[TF142]
                                            ,[TF143]
                                            ,[TF144]
                                            ,[TF145]
                                            ,[TF146]
                                            ,[TF147]
                                            ,[TF148]
                                            ,[TF149]
                                            ,[TF150]
                                            ,[TF151]
                                            ,[TF152]
                                            ,[TF153]
                                            ,[TF154]
                                            ,[TF155]
                                            ,[TF156]
                                            ,[TF157]
                                            ,[TF158]
                                            ,[TF159]
                                            ,[TF160]
                                            ,[TF161]
                                            ,[TF162]
                                            ,[TF163]
                                            ,[TF164]
                                            ,[TF165]
                                            ,[TF166]
                                            ,[TF167]
                                            ,[TF168]
                                            ,[TF169]
                                            ,[TF170]
                                            ,[TF171]
                                            ,[TF172]
                                            ,[TF173]
                                            ,[UDF01]
                                            ,[UDF02]
                                            ,[UDF03]
                                            ,[UDF04]
                                            ,[UDF05]
                                            ,[UDF06]
                                            ,[UDF07]
                                            ,[UDF08]
                                            ,[UDF09]
                                            ,[UDF10]
                                            )

                                            SELECT 
                                            PURTD.[COMPANY]
                                            ,PURTD.[CREATOR] AS [CREATOR]
                                            ,PURTD.[USR_GROUP] AS [USR_GROUP]
                                            ,PURTD.[CREATE_DATE] AS [CREATE_DATE]
                                            ,PURTD.[MODIFIER] AS [MODIFIER]
                                            ,PURTD.[MODI_DATE] AS [MODI_DATE]
                                            ,PURTD.[FLAG] AS [FLAG]
                                            ,PURTD.[CREATE_TIME] AS [CREATE_TIME]
                                            ,PURTD.[MODI_TIME] AS [MODI_TIME]
                                            ,PURTD.[TRANS_TYPE] AS [TRANS_TYPE]
                                            ,'PURI08' AS [TRANS_NAME]
                                            ,PURTD.[sync_date] AS [sync_date]
                                            ,PURTD.[sync_time] AS [sync_time]
                                            ,PURTD.[sync_mark] AS [sync_mark]
                                            ,PURTD.[sync_count] AS [sync_count]
                                            ,PURTD.[DataUser] AS [DataUser]
                                            ,PURTD.[DataGroup] AS [DataGroup]
                                            ,TD001 AS [TF001]
                                            ,TD002 AS [TF002]
                                            ,'{3}' AS [TF003]
                                            ,TD003 AS [TF004]
                                            ,[PURTATBCHAGE].TB004 AS [TF005]
                                            ,[PURTATBCHAGE].TB005 AS [TF006]
                                            ,[PURTATBCHAGE].TB006 AS [TF007]
                                            ,TD007 AS [TF008]
                                            ,[PURTATBCHAGE].TB009 AS [TF009]
                                            ,TD009 AS [TF010]
                                            ,TD010 AS [TF011]
                                            ,[PURTATBCHAGE].TB009*TD010 AS [TF012]
                                            ,[PURTATBCHAGE].TB011 AS [TF013]
                                            ,'N' AS [TF014]
                                            ,TD015 AS [TF015]
                                            ,'N' AS [TF016]
                                            ,[PURTATBCHAGE].TB012 AS [TF017]
                                            ,TD019 AS [TF018]
                                            ,TD020 AS [TF019]
                                            ,TD022 AS [TF020]
                                            ,TD025 AS [TF021]
                                            ,TD017 AS [TF022]
                                            ,TD029 AS [TF023]
                                            ,TD030 AS [TF024]
                                            ,TD032 AS [TF025]
                                            ,TD033 AS [TF026]
                                            ,TD036 AS [TF027]
                                            ,TD037 AS [TF028]
                                            ,TD038 AS [TF029]
                                            ,TD014 AS [TF030]
                                            ,'' AS [TF031]
                                            ,'' AS [TF032]
                                            ,'' AS [TF033]
                                            ,'' AS [TF034]
                                            ,'' AS [TF035]
                                            ,0 AS [TF036]
                                            ,0 AS [TF037]
                                            ,'' AS [TF038]
                                            ,'' AS [TF039]
                                            ,'' AS [TF040]
                                            ,0 AS [TF041]
                                            ,TD003 AS [TF104]
                                            ,TD004 AS [TF105]
                                            ,TD005 AS [TF106]
                                            ,TD006 AS [TF107]
                                            ,TD007 AS [TF108]
                                            ,TD008 AS [TF109]
                                            ,TD009 AS [TF110]
                                            ,TD010 AS [TF111]
                                            ,TD011 AS [TF112]
                                            ,TD012 AS [TF113]
                                            ,TD016 AS [TF114]
                                            ,TD019 AS [TF118]
                                            ,TD020 AS [TF119]
                                            ,TD022 AS [TF120]
                                            ,TD025 AS [TF121]
                                            ,TD017 AS [TF122]
                                            ,TD029 AS [TF123]
                                            ,TD030 AS [TF124]
                                            ,TD031 AS [TF125]
                                            ,TD032 AS [TF126]
                                            ,TD033 AS [TF127]
                                            ,TD034 AS [TF128]
                                            ,TD035 AS [TF129]
                                            ,TD034 AS [TF130]
                                            ,TD035 AS [TF131]
                                            ,TD036 AS [TF132]
                                            ,TD037 AS [TF133]
                                            ,TD038 AS [TF134]
                                            ,TD014 AS [TF135]
                                            ,0 AS [TF136]
                                            ,0 AS [TF137]
                                            ,'' AS [TF138]
                                            ,'' AS [TF139]
                                            ,'' AS [TF140]
                                            ,0 AS [TF141]
                                            ,'' AS [TF142]
                                            ,'' AS [TF143]
                                            ,'' AS [TF144]
                                            ,'2' AS [TF145]
                                            ,'2' AS [TF146]
                                            ,'' AS [TF147]
                                            ,'' AS [TF148]
                                            ,'' AS [TF149]
                                            ,'' AS [TF150]
                                            ,'' AS [TF151]
                                            ,TD080 AS [TF152]
                                            ,TD081 AS [TF153]
                                            ,TD082 AS [TF154]
                                            ,TD083 AS [TF155]
                                            ,TD080 AS [TF156]
                                            ,TD081 AS [TF157]
                                            ,TD082 AS [TF158]
                                            ,TD083 AS [TF159]
                                            ,TD084 AS [TF160]
                                            ,TD085 AS [TF161]
                                            ,TD084 AS [TF162]
                                            ,TD085 AS [TF163]
                                            ,0 AS [TF164]
                                            ,0 AS [TF165]
                                            ,0 AS [TF166]
                                            ,0 AS [TF167]
                                            ,'' AS [TF168]
                                            ,'' AS [TF169]
                                            ,'' AS [TF170]
                                            ,'' AS [TF171]
                                            ,'' AS [TF172]
                                            ,'' AS [TF173]
                                            ,CONVERT(NVARCHAR,[PURTATBCHAGE].VERSIONS)+CONVERT(NVARCHAR,[PURTATBCHAGE].TA001)+CONVERT(NVARCHAR,[PURTATBCHAGE].TA002)+CONVERT(NVARCHAR,[PURTATBCHAGE].TB003) AS [UDF01]
                                            ,'' AS [UDF02]
                                            ,'' AS [UDF03]
                                            ,'' AS [UDF04]
                                            ,'' AS [UDF05]
                                            ,0 AS [UDF06]
                                            ,0 AS [UDF07]
                                            ,0 AS [UDF08]
                                            ,0 AS [UDF09]
                                            ,0 AS [UDF10]

                                            FROM [TK].dbo.PURTD,[TKPUR].[dbo].[PURTATBCHAGE]
                                            WHERE 1=1
                                            AND PURTD.TD026=[PURTATBCHAGE].TA001 AND PURTD.TD027=[PURTATBCHAGE].TA002 AND PURTD.TD028=[PURTATBCHAGE].TB003
                                            AND TD001='{4}' AND TD002='{5}'
                                            AND [PURTATBCHAGE].TA001='{0}' AND [PURTATBCHAGE].TA002='{1}' AND [PURTATBCHAGE].VERSIONS='{2}'



                                            INSERT INTO [TK].[dbo].[PURTE]
                                            (
                                            [COMPANY]
                                            ,[CREATOR]
                                            ,[USR_GROUP]
                                            ,[CREATE_DATE]
                                            ,[MODIFIER]
                                            ,[MODI_DATE]
                                            ,[FLAG]
                                            ,[CREATE_TIME]
                                            ,[MODI_TIME]
                                            ,[TRANS_TYPE]
                                            ,[TRANS_NAME]
                                            ,[sync_date]
                                            ,[sync_time]
                                            ,[sync_mark]
                                            ,[sync_count]
                                            ,[DataUser]
                                            ,[DataGroup]
                                            ,[TE001]
                                            ,[TE002]
                                            ,[TE003]
                                            ,[TE004]
                                            ,[TE005]
                                            ,[TE006]
                                            ,[TE007]
                                            ,[TE008]
                                            ,[TE009]
                                            ,[TE010]
                                            ,[TE011]
                                            ,[TE012]
                                            ,[TE013]
                                            ,[TE014]
                                            ,[TE015]
                                            ,[TE016]
                                            ,[TE017]
                                            ,[TE018]
                                            ,[TE019]
                                            ,[TE020]
                                            ,[TE021]
                                            ,[TE022]
                                            ,[TE023]
                                            ,[TE024]
                                            ,[TE025]
                                            ,[TE026]
                                            ,[TE027]
                                            ,[TE028]
                                            ,[TE029]
                                            ,[TE030]
                                            ,[TE031]
                                            ,[TE032]
                                            ,[TE033]
                                            ,[TE034]
                                            ,[TE035]
                                            ,[TE036]
                                            ,[TE037]
                                            ,[TE038]
                                            ,[TE039]
                                            ,[TE040]
                                            ,[TE041]
                                            ,[TE042]
                                            ,[TE043]
                                            ,[TE045]
                                            ,[TE046]
                                            ,[TE047]
                                            ,[TE048]
                                            ,[TE103]
                                            ,[TE107]
                                            ,[TE108]
                                            ,[TE109]
                                            ,[TE110]
                                            ,[TE113]
                                            ,[TE114]
                                            ,[TE115]
                                            ,[TE118]
                                            ,[TE119]
                                            ,[TE120]
                                            ,[TE121]
                                            ,[TE122]
                                            ,[TE123]
                                            ,[TE124]
                                            ,[TE125]
                                            ,[TE134]
                                            ,[TE135]
                                            ,[TE136]
                                            ,[TE137]
                                            ,[TE138]
                                            ,[TE139]
                                            ,[TE140]
                                            ,[TE141]
                                            ,[TE142]
                                            ,[TE143]
                                            ,[TE144]
                                            ,[TE145]
                                            ,[TE146]
                                            ,[TE147]
                                            ,[TE148]
                                            ,[TE149]
                                            ,[TE150]
                                            ,[TE151]
                                            ,[TE152]
                                            ,[TE153]
                                            ,[TE154]
                                            ,[TE155]
                                            ,[TE156]
                                            ,[TE157]
                                            ,[TE158]
                                            ,[TE159]
                                            ,[TE160]
                                            ,[TE161]
                                            ,[TE162]
                                            ,[UDF01]
                                            ,[UDF02]
                                            ,[UDF03]
                                            ,[UDF04]
                                            ,[UDF05]
                                            ,[UDF06]
                                            ,[UDF07]
                                            ,[UDF08]
                                            ,[UDF09]
                                            ,[UDF10]
                                            )
                                            SELECT 
                                            PURTC.[COMPANY]
                                            ,PURTC.[CREATOR]
                                            ,PURTC.[USR_GROUP]
                                            ,PURTC.[CREATE_DATE]
                                            ,PURTC.[MODIFIER]
                                            ,PURTC.[MODI_DATE]
                                            ,PURTC.[FLAG]
                                            ,PURTC.[CREATE_TIME]
                                            ,PURTC.[MODI_TIME]
                                            ,PURTC.[TRANS_TYPE]
                                            ,'PURI08' AS [TRANS_NAME]
                                            ,PURTC.[sync_date]
                                            ,PURTC.[sync_time]
                                            ,PURTC.[sync_mark]
                                            ,PURTC.[sync_count]
                                            ,PURTC.[DataUser]
                                            ,PURTC.[DataGroup]
                                            ,TC001 AS [TE001]
                                            ,TC002 AS [TE002]
                                            ,'{3}' AS [TE003]
                                            ,CONVERT(NVARCHAR,GETDATE(),112) AS [TE004]
                                            ,TC004 AS [TE005]
                                            ,'' AS [TE006]
                                            ,TC005 AS [TE007]
                                            ,TC006 AS [TE008]
                                            ,TC007 AS [TE009]
                                            ,TC008 AS [TE010]
                                            ,CONVERT(NVARCHAR,GETDATE(),112) AS [TE011]
                                            ,'N' AS [TE012]
                                            ,TC015 AS [TE013]
                                            ,TC016 AS [TE014]
                                            ,TC017 AS [TE015]
                                            ,0 AS [TE016]
                                            ,'N' AS [TE017]
                                            ,TC018 AS [TE018]
                                            ,TC021 AS [TE019]
                                            ,TC022 AS [TE020]
                                            ,'' AS [TE021]
                                            ,TC026 AS [TE022]
                                            ,TC027 AS [TE023]
                                            ,TC028 AS [TE024]
                                            ,'N' AS [TE025]
                                            ,0 AS [TE026]
                                            ,TC009 AS [TE027]
                                            ,'N' AS [TE028]
                                            ,TC035 AS [TE029]
                                            ,'' AS [TE030]
                                            ,'' AS [TE031]
                                            ,'N' AS [TE032]
                                            ,'' AS [TE033]
                                            ,0 AS [TE034]
                                            ,0 AS [TE035]
                                            ,'' AS [TE036]
                                            ,TC011 AS [TE037]
                                            ,'' AS [TE038]
                                            ,'' AS [TE039]
                                            ,'' AS [TE040]
                                            ,TC050 AS [TE041]
                                            ,'' AS [TE042]
                                            ,TC036 AS [TE043]
                                            ,TC037 AS [TE045]
                                            ,TC038 AS [TE046]
                                            ,TC039 AS [TE047]
                                            ,TC040 AS [TE048]
                                            ,'' AS [TE103]
                                            ,TC005 AS [TE107]
                                            ,TC006 AS [TE108]
                                            ,TC007 AS [TE109]
                                            ,TC008 AS [TE110]
                                            ,TC015 AS [TE113]
                                            ,TC016 AS [TE114]
                                            ,TC017 AS [TE115]
                                            ,TC018 AS [TE118]
                                            ,TC021 AS [TE119]
                                            ,TC022 AS [TE120]
                                            ,TC026 AS [TE121]
                                            ,TC027 AS [TE122]
                                            ,TC028 AS [TE123]
                                            ,TC009 AS [TE124]
                                            ,TC035 AS [TE125]
                                            ,0 AS [TE134]
                                            ,0 AS [TE135]
                                            ,'' AS [TE136]
                                            ,'' AS [TE137]
                                            ,'' AS [TE138]
                                            ,'' AS [TE139]
                                            ,'1' AS [TE140]
                                            ,'N' AS [TE141]
                                            ,'N' AS [TE142]
                                            ,TC036 AS [TE143]
                                            ,'N' AS [TE144]
                                            ,'' AS [TE145]
                                            ,TC041 AS [TE146]
                                            ,TC041 AS [TE147]
                                            ,TC011 AS [TE148]
                                            ,0 AS [TE149]
                                            ,0 AS [TE150]
                                            ,0 AS [TE151]
                                            ,0 AS [TE152]
                                            ,'' AS [TE153]
                                            ,'' AS [TE154]
                                            ,'' AS [TE155]
                                            ,'' AS [TE156]
                                            ,'' AS [TE157]
                                            ,'' AS [TE158]
                                            ,TC037 AS [TE159]
                                            ,TC038 AS [TE160]
                                            ,TC039 AS [TE161]
                                            ,TC040 AS [TE162]
                                            ,'' AS [UDF01]
                                            ,'' AS [UDF02]
                                            ,'' AS [UDF03]
                                            ,'' AS [UDF04]
                                            ,'' AS [UDF05]
                                            ,0 AS [UDF06]
                                            ,0 AS [UDF07]
                                            ,0 AS [UDF08]
                                            ,0 AS [UDF09]
                                            ,0 AS [UDF10]
                                            FROM  [TK].dbo.PURTC
                                            WHERE 1=1
                                            AND TC001='{4}' AND TC002='{5}'
                                            ", DR["TA001"].ToString(), DR["TA002"].ToString(), DR["VERSIONS"].ToString(), DR["TE003"].ToString(), DR["TE001"].ToString(), DR["TE002"].ToString());
                    }

                    sbSql.AppendFormat(@"  
                                   
                                        ");

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
        }

        public void SEARCHPURTE(string TA001, string TA002,string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                 
                                    SELECT TE001 AS '採購變更單別',TE002 AS '採購變更單號',TE003 AS '版次'
                                    FROM [TK].dbo.PURTE
                                    WHERE TE001+TE002 IN 
                                    (
                                    SELECT TD001+TD002
                                    FROM [TK].dbo.PURTD
                                    WHERE TD026+TD027+TD028 IN 
                                    (
                                    SELECT TA001+TA002+TB003
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE  TA001='{0}' AND TA002='{1}' AND VERSIONS='{2}'
                                    )
                                    GROUP BY  TD001,TD002
                                    )

                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView6.DataSource = ds.Tables["ds"];
                        dataGridView6.AutoResizeColumns();

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
            string ISUOFSTATUS = "N";
            string DOC_NBR = "";

            ISUOFSTATUS = SERACHlDOC_NBR_CHECK(textBox11.Text + textBox12.Text + textBox10.Text);
            textBox26.Text = SERACHlDOC_NBR(textBox11.Text + textBox12.Text + textBox10.Text);
            DOC_NBR = SERACHlDOC_NBR(textBox11.Text + textBox12.Text + textBox10.Text);
            

            //沒有產生過UOF表單
            if (string.IsNullOrEmpty(DOC_NBR))
            {
                ADDTB_WKF_EXTERNAL_TASK(textBox11.Text.Trim(), textBox12.Text.Trim(), textBox10.Text.Trim());
            }
            //UOF表單是否沒有在簽核中
            else if(!string.IsNullOrEmpty(DOC_NBR) && ISUOFSTATUS.Equals("N"))
            {
                ADDTB_WKF_EXTERNAL_TASK(textBox11.Text.Trim(), textBox12.Text.Trim(), textBox10.Text.Trim());
            }
            else
            {
                MessageBox.Show(textBox11.Text + textBox12.Text + textBox10.Text+"已發出請購變更單"+ DOC_NBR + " ，但未完成簽核。");

            }

            
        }
        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                DELPURTATBCHAGE(textBox10.Text.Trim(), textBox11.Text.Trim(), textBox12.Text.Trim(), textBox14.Text.Trim());
                SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            SEARCHPURTATBCHAGEDETAILS(textBox5.Text, textBox6.Text);
        }
        private void button10_Click(object sender, EventArgs e)
        {
            NEWPURTEPURTF(textBox27.Text, textBox28.Text, textBox29.Text);
        }

        #endregion


    }
}
