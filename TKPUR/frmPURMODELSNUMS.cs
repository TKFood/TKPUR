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
using System.Globalization;

namespace TKPUR
{
    public partial class frmPURMODELSNUMS : Form
    {

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;

        public frmPURMODELSNUMS()
        {
            InitializeComponent();

            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            // 取得今年的第一天
            DateTime firstDayOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            // 取得今年的最後一天
            DateTime lastDayOfYear = new DateTime(DateTime.Now.Year, 12, 31);

            dateTimePicker1.Value = firstDayOfYear;
            dateTimePicker2.Value = lastDayOfYear;
        }
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT  [ID],[KIND],[PARAID],[PARANAME] FROM [TKPUR].[dbo].[TBPARA] WHERE [KIND]='是否結案' ORDER BY ID
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARANAME";
            comboBox1.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT  [ID],[KIND],[PARAID],[PARANAME] FROM [TKPUR].[dbo].[TBPARA] WHERE [KIND]='是否結案2' ORDER BY ID
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARANAME";
            comboBox2.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='PAYKINDS' 
                                ORDER BY [PARANAME]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "PARAID";
            comboBox3.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void comboBox4load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='PAYKINDS' 
                                AND [PARAID] NOT IN ('全部')
                                ORDER BY [PARANAME]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "PARAID";
            comboBox4.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void SEARCH_PURMODELSNUMS(string NAMES, string MB001, string ISCLOSE, string PAYKINDS, string SDAYS, string EDAYS, string COMMENTS, string MB002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();
            StringBuilder sbSqlQuery4 = new StringBuilder();
            StringBuilder sbSqlQuery5 = new StringBuilder();
            StringBuilder sbSqlQuery6 = new StringBuilder();
            StringBuilder sbSqlQuery7 = new StringBuilder();
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

                if (!string.IsNullOrEmpty(NAMES))
                {
                    sbSqlQuery1.AppendFormat(@" AND [NAMES] LIKE '%{0}%'", NAMES);
                }
                else
                {
                    sbSqlQuery1.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSqlQuery2.AppendFormat(@" AND [MB001] LIKE '%{0}%'", MB001);
                }
                else
                {
                    sbSqlQuery2.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(ISCLOSE) && ISCLOSE.Equals("全部"))
                {
                    sbSqlQuery3.AppendFormat(@"");
                }
                else if (!string.IsNullOrEmpty(ISCLOSE))
                {
                    sbSqlQuery3.AppendFormat(@" AND [ISCLOSE] LIKE '%{0}%'", ISCLOSE);
                }
                if (!string.IsNullOrEmpty(PAYKINDS) && PAYKINDS.Equals("全部"))
                {
                    sbSqlQuery4.AppendFormat(@" ");
                }
                else if (!string.IsNullOrEmpty(PAYKINDS))
                {
                    sbSqlQuery4.AppendFormat(@" AND [PAYKINDS] IN ('{0}')", PAYKINDS);
                }
                else
                {
                    sbSqlQuery4.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(SDAYS) && !string.IsNullOrEmpty(EDAYS))
                {
                    sbSqlQuery5.AppendFormat(@" AND CONVERT(NVARCHAR,[CREATEDATES],112)>='{0}' AND CONVERT(NVARCHAR,[CREATEDATES],112)<='{1}'", SDAYS, EDAYS);
                }

                if (!string.IsNullOrEmpty(COMMENTS))
                {
                    sbSqlQuery6.AppendFormat(@" AND [COMMENTS] LIKE '%{0}%'", COMMENTS);
                }
                else
                {
                    sbSqlQuery6.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(MB002))
                {
                    sbSqlQuery7.AppendFormat(@" AND [MB002] LIKE '%{0}%'", MB002);
                }
                else
                {
                    sbSqlQuery7.AppendFormat(@" ");
                }

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT 
                                     [NAMES] AS '版型' 
                                    ,[MB001] AS '品號' 
                                    ,[MB002] AS '品名' 
                                    ,[BACKMONEYS] AS '可退還的版費' 
                                    ,[TARGETNUMS] AS '目標進貨量' 
                                    ,[TOTALNUMS] AS '已進貨量' 
                                    ,[ISCLOSE] AS '是否結案' 
                                    ,[PAYKINDS] AS '付款別'
                                    ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '建立日期'
                                    ,[COMMENTS] AS '備註'
                                    FROM [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    {2}
                                    {3}
                                    {4}
                                    {5}
                                    {6}
                                    ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString(), sbSqlQuery4.ToString(), sbSqlQuery5.ToString(), sbSqlQuery6.ToString(), sbSqlQuery7.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds.Tables["ds"];
                    dataGridView1.AutoResizeColumns();
                    // 設定券消費列的數字格式
                    dataGridView1.Columns["已進貨量"].DefaultCellStyle.Format = "#,##0";
                    dataGridView1.Columns["已進貨量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns["可退還的版費"].DefaultCellStyle.Format = "#,##0";
                    dataGridView1.Columns["可退還的版費"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    dataGridView1.Columns["目標進貨量"].DefaultCellStyle.Format = "#,##0";
                    dataGridView1.Columns["目標進貨量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                }
                else
                {
                    dataGridView1.DataSource = null;

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
            DateTime CREATEDATES;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox3.Text = row.Cells["版型"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["可退還的版費"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["目標進貨量"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["已進貨量"].Value.ToString().Trim();
                    textBox10.Text = row.Cells["備註"].Value.ToString().Trim();

                    comboBox2.Text = row.Cells["是否結案"].Value.ToString().Trim();
                    comboBox4.Text = row.Cells["付款別"].Value.ToString().Trim();

                    DateTime.TryParseExact(row.Cells["建立日期"].Value.ToString().Trim(), "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out CREATEDATES);
                    dateTimePicker3.Value = CREATEDATES;

                }
                else
                {
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox10.Text = "";
                }


            }
        }

        public void ADD_PURVERSIONSNUMS(string NAMES, string MB001, string MB002, string BACKMONEYS, string TARGETNUMS, string TOTALNUMS, string ISCLOSE, string PAYKINDS, string CREATEDATES, string COMMENTS)
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

                TOTALNUMS = "0";
                CREATEDATES = DateTime.Now.ToString("yyyy/MM/dd");
                sbSql.AppendFormat(@"  
                                   INSERT INTO [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    (NAMES,MB001,MB002,BACKMONEYS,TARGETNUMS,TOTALNUMS,ISCLOSE,PAYKINDS,CREATEDATES,COMMENTS)
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')
                                    ", NAMES, MB001, MB002, BACKMONEYS, TARGETNUMS, TOTALNUMS, ISCLOSE, PAYKINDS, CREATEDATES, COMMENTS);

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
        public void UPDATE_PURVERSIONSNUMS(string NAMES, string MB001, string MB002, string BACKMONEYS, string TARGETNUMS, string TOTALNUMS, string ISCLOSE, string PAYKINDS, string CREATEDATES, string COMMENTS)
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
                                   
                                    UPDATE  [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    SET MB001='{1}',MB002='{2}',BACKMONEYS='{3}',TARGETNUMS='{4}',TOTALNUMS='{5}',ISCLOSE='{6}',PAYKINDS='{7}',CREATEDATES='{8}',COMMENTS='{9}'
                                    WHERE NAMES='{0}'
                                    ", NAMES, MB001, MB002, BACKMONEYS, TARGETNUMS, TOTALNUMS, ISCLOSE, PAYKINDS, CREATEDATES, COMMENTS);

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
        public void DELETE_PURVERSIONSNUMS(string NAMES)
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
                                    DELETE  [TKPUR].[dbo].[PURVERSIONSNUMS]                                    
                                    WHERE NAMES='{0}'
                                    ", NAMES);

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
        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            DataTable DT = FINDPURTE(textBox4.Text.Trim());
            if (DT != null && DT.Rows.Count >= 1)
            {
                textBox5.Text = DT.Rows[0]["MB002"].ToString();
            }
            else
            {
                textBox5.Text = "";
            }
        }
        public DataTable FINDPURTE(string MB001)
        {
            DataTable DT = new DataTable();

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
                                    SELECT MB002 
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001='{0}'
                                             ", MB001);

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
        /// <summary>
        /// 更新已進貨的數量，用驗收數量>TOTALNUMS
        /// </summary>
        public void UDPATE_PURVERSIONSNUMS_TOTALNUMS()
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
                                    UPDATE [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    SET [TOTALNUMS]=(SELECT SUM(TH015) FROM[TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TG013='Y' AND TH004=MB001 ) 
                                    WHERE [TOTALNUMS]<>(SELECT SUM(TH015) FROM[TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TG013='Y' AND TH004=MB001 ) 

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

        public void SETFASTREPORT(string NAMES, string MB001, string ISCLOSE, string PAYKINDS, string SDAYS, string EDAYS, string COMMENTS, string MB002)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();
            StringBuilder sbSqlQuery4 = new StringBuilder();
            StringBuilder sbSqlQuery5 = new StringBuilder();
            StringBuilder sbSqlQuery6 = new StringBuilder();
            StringBuilder sbSqlQuery7 = new StringBuilder();

            if (!string.IsNullOrEmpty(NAMES))
            {
                sbSqlQuery1.AppendFormat(@" AND [NAMES] LIKE '%{0}%'", NAMES);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(MB001))
            {
                sbSqlQuery2.AppendFormat(@" AND [MB001] LIKE '%{0}%'", MB001);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(ISCLOSE) && ISCLOSE.Equals("全部"))
            {
                sbSqlQuery3.AppendFormat(@"");
            }
            else if (!string.IsNullOrEmpty(ISCLOSE))
            {
                sbSqlQuery3.AppendFormat(@" AND [ISCLOSE] LIKE '%{0}%'", ISCLOSE);
            }
            if (!string.IsNullOrEmpty(PAYKINDS) && PAYKINDS.Equals("全部"))
            {
                sbSqlQuery4.AppendFormat(@" ");
            }
            else if (!string.IsNullOrEmpty(PAYKINDS))
            {
                sbSqlQuery4.AppendFormat(@" AND [PAYKINDS] IN ('{0}')", PAYKINDS);
            }
            else
            {
                sbSqlQuery4.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(SDAYS) && !string.IsNullOrEmpty(EDAYS))
            {
                sbSqlQuery5.AppendFormat(@" AND CONVERT(NVARCHAR,[CREATEDATES],112)>='{0}' AND CONVERT(NVARCHAR,[CREATEDATES],112)<='{1}'", SDAYS, EDAYS);
            }

            if (!string.IsNullOrEmpty(COMMENTS))
            {
                sbSqlQuery6.AppendFormat(@" AND [COMMENTS] LIKE '%{0}%'", COMMENTS);
            }
            else
            {
                sbSqlQuery6.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(MB002))
            {
                sbSqlQuery7.AppendFormat(@" AND [MB002] LIKE '%{0}%'", MB002);
            }
            else
            {
                sbSqlQuery7.AppendFormat(@" ");
            }

            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@"
                                    SELECT 
                                     [NAMES] AS '版型' 
                                    ,[MB001] AS '品號' 
                                    ,[MB002] AS '品名' 
                                    ,[BACKMONEYS] AS '可退還的版費' 
                                    ,[TARGETNUMS] AS '目標進貨量' 
                                    ,[TOTALNUMS] AS '已進貨量' 
                                    ,[ISCLOSE] AS '是否結案' 
                                    ,[PAYKINDS] AS '付款別'
                                    ,CONVERT(NVARCHAR,[CREATEDATES],112) AS '建立日期'
                                    ,[COMMENTS] AS '備註'
                                    FROM [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    {2}
                                    {3}
                                    {4}
                                    {5}
                                    {6}
                                    ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString(), sbSqlQuery4.ToString(), sbSqlQuery5.ToString(), sbSqlQuery6.ToString(), sbSqlQuery7.ToString());
            SQL1 = sbSql;

            Report report1 = new Report();
            report1.Load(@"REPORT\版費.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.Preview = previewControl1;
            report1.Show();
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_PURMODELSNUMS(textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString(), comboBox3.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox9.Text.Trim(), textBox11.Text.Trim());
            SETFASTREPORT(textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString(), comboBox3.Text.ToString(), dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox9.Text.Trim(), textBox11.Text.Trim());
        }

        #endregion
    }
}
