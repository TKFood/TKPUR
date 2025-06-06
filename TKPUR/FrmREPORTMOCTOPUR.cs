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
    public partial class FrmREPORTMOCTOPUR : Form
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
        public Report report1 { get; private set; }

        public FrmREPORTMOCTOPUR()
        {
            InitializeComponent();
        }

        public class DATA_SET
        {
            public string COMPANY;
            public string CREATOR;
            public string USR_GROUP;
            public string CREATE_DATE;
            public string MODIFIER;
            public string MODI_DATE;
            public string FLAG;
            public string CREATE_TIME;
            public string MODI_TIME;
            public string TRANS_TYPE;
            public string TRANS_NAME;
            public string sync_date;
            public string sync_time;
            public string sync_mark;
            public string sync_count;
            public string DataUser;
            public string DataGroup;


        }

        #region FUNCTION

        #endregion
        private void frnREPORTMOCTOPUR_Load(object sender, EventArgs e)
        {
            SETDATES();
            SET_TEXT();
        }

        public void SETDATES()
        {
            // 取得今年的第一天
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            // 取得今天
            DateTime lastDayOfMonth = DateTime.Now;

            // 取得當前日期
            DateTime today = DateTime.Today;
            // 取得當前月份的第一天，並加一個月
            DateTime firstDayOfNextMonth = new DateTime(today.Year, today.Month, 1).AddMonths(1);
            // 減去一天即為本月的最後一天
            DateTime lastDayOfThisMonth = firstDayOfNextMonth.AddDays(-1);

            dateTimePicker1.Value = firstDayOfMonth;
            dateTimePicker2.Value = lastDayOfThisMonth;
            dateTimePicker3.Value = firstDayOfMonth;
            dateTimePicker4.Value = lastDayOfThisMonth;
            dateTimePicker5.Value = firstDayOfMonth;
            dateTimePicker6.Value = lastDayOfThisMonth;
            dateTimePicker7.Value = firstDayOfMonth;
            dateTimePicker8.Value = lastDayOfThisMonth;

            dateTimePicker10.Value = firstDayOfMonth;
            dateTimePicker11.Value = lastDayOfThisMonth;
        }

        public void SET_TEXT()
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
                                WHERE [KIND] ='FrmREPORTMOCTOPUR'
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("KIND", typeof(string));
            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);         
            sqlConn.Close();

            foreach(DataRow DR in dt.Rows)
            {
                if(DR["PARAID"].ToString().Equals("公司電話"))
                {
                    textBox4.Text = DR["PARANAME"].ToString();
                    textBox11.Text = DR["PARANAME"].ToString();
                }
                else if(DR["PARAID"].ToString().Equals("公司傳真"))
                {
                    textBox5.Text = DR["PARANAME"].ToString();
                    textBox12.Text = DR["PARANAME"].ToString();
                }
                else if(DR["PARAID"].ToString().Equals("送貨地址"))
                {
                    textBox6.Text = DR["PARANAME"].ToString();
                    textBox13.Text = DR["PARANAME"].ToString();
                }
                else if(DR["PARAID"].ToString().Equals("營業稅率"))
                {
                    textBox7.Text = DR["PARANAME"].ToString();
                    textBox14.Text = DR["PARANAME"].ToString();
                }
            }

        }
        public void SEARCH(string TA001,string SDAYS,string EDAYS)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
                                    TA001 AS '單別'
                                    ,TA002 AS '單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TA003),111) AS '單據日期'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TA010),111) AS '到貨日'
                                    ,TA032 AS '廠商代號'
                                    ,TA023 AS '單位'
                                    ,TA006 AS '品號'
                                    ,TA034 AS '品名'
                                    ,TA035 AS '規格'
                                    ,TA022 AS '採購單價'
                                    ,TA015 AS '採購數量'
                                    ,TA042 AS '交易幣別'
                                    ,TA043 AS '匯率'
                                    ,TA019 AS '廠別代號'
                                    ,TA020 AS '交貨庫別'
                                    ,TA029 AS '備註'
                                    ,MA002 AS '廠商'
                                    ,MA003 AS '廠商全名'
                                    ,MA008 AS '廠商電話'
                                    ,MA013 AS '聯絡人'
                                    ,MA055 AS '付款條件'
                                    ,MA025 AS '付款'
                                    --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅  &&880210 &&88-11-25 OLD:預留C10
                                    ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '課稅別'
                                    ,MA047 AS '採購人員'
                                    ,MA010 AS '廠商傳真'
                                    ,MV002 AS '採購人'
                                    ,(TA022*TA015) AS '採購金額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '稅額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)+(TA022*TA015)*0.05) WHEN MA044=3 THEN (TA022*TA015) WHEN MA044=4 THEN (TA022*TA015) WHEN MA044=9 THEN (TA022*TA015) END )  AS '金額合計'

                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TA032
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=MA047
                                    WHERE TA001='{0}'
                                    AND TA003>='{1}' AND TA003<='{2}'
                                    ", TA001,SDAYS,EDAYS);



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["ds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                        // 設置欄位順序
                        dataGridView1.Columns["單別"].DisplayIndex = 0;
                        dataGridView1.Columns["單號"].DisplayIndex = 1;
                        dataGridView1.Columns["廠商"].DisplayIndex = 2;
                        dataGridView1.Columns["品名"].DisplayIndex = 3;
                        dataGridView1.Columns["採購數量"].DisplayIndex = 4;


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
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox2.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox3.Text = row.Cells["單號"].Value.ToString().Trim();
                }
            }
        }

        public void SETFASTREPORT(string TA001,string TA002,string P1,string P2,string P3,string P4)
        {
            report1 = new Report();
            string SQL = "";
            
            report1.Load(@"REPORT\託外採購單.frx");
            
            SQL = SETFASETSQL1(TA001, TA002);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
           

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;

            report1.SetParameterValue("公司電話", P1);
            report1.SetParameterValue("製表日期", DateTime.Now.ToString("yyyy/MM/dd"));
            report1.SetParameterValue("公司傳真", P2);
            report1.SetParameterValue("送貨地址", P3);
            report1.SetParameterValue("營業稅率", P4);

            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();

        }
        public string SETFASETSQL1(string TA001, string TA002)
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"   
                                SELECT 
                                TA001 AS '單別'
                                ,TA002 AS '單號'
                                ,CONVERT(NVARCHAR,CONVERT(datetime,TA003),111) AS '單據日期'
                                ,CONVERT(NVARCHAR,CONVERT(datetime,TA010),111) AS '到貨日'
                                ,TA032 AS '廠商代號'
                                ,TA023 AS '單位'
                                ,TA006 AS '品號'
                                ,TA034 AS '品名'
                                ,TA035 AS '規格'
                                ,TA022 AS '採購單價'
                                ,TA015 AS '採購數量'
                                ,TA042 AS '交易幣別'
                                ,TA043 AS '匯率'
                                ,TA019 AS '廠別代號'
                                ,TA020 AS '交貨庫別'
                                ,TA029 AS '備註'
                                ,MA002 AS '廠商'
                                ,MA003 AS '廠商全名'
                                ,MA008 AS '廠商電話'
                                ,MA013 AS '聯絡人'
                                ,MA055 AS '付款條件'
                                ,MA025 AS '付款'
                                --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅  &&880210 &&88-11-25 OLD:預留C10
                                ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '課稅別'
                                ,MA047 AS '採購人員'
                                ,MA010 AS '廠商傳真'
                                ,MV002 AS '採購人'
                                ,(TA022*TA015) AS '採購金額'
                                ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '稅額'
                                ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)+(TA022*TA015)*0.05) WHEN MA044=3 THEN (TA022*TA015) WHEN MA044=4 THEN (TA022*TA015) WHEN MA044=9 THEN (TA022*TA015) END )  AS '金額合計'

                                FROM [TK].dbo.MOCTA
                                LEFT JOIN [TK].dbo.PURMA ON MA001=TA032
                                LEFT JOIN [TK].dbo.CMSMV ON MV001=MA047
                                WHERE 1=1
                                AND TA001='{0}'
                                AND TA002='{1}'                            

                                ", TA001, TA002);

            return FASTSQL.ToString();
        }

        public void SEARCH_MOCTO(string TA001, string SDAYS, string EDAYS)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
                                    TO001 AS '單別'
                                    ,TO002 AS '單號'
                                    ,TO003 AS '變更版次'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TO004),111) AS '單據日期'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TO013),111) AS '到貨日'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TO012),111) AS '採購日期'
                                    ,TO033 AS '廠商代號'
                                    ,TO010 AS '單位'
                                    ,TO009 AS '品號'
                                    ,TO035 AS '品名'
                                    ,TO036 AS '規格'
                                    ,TO024 AS '採購單價'
                                    ,TO017 AS '採購數量'
                                    ,TO045 AS '交易幣別'
                                    ,TO046 AS '匯率'
                                    ,TO021 AS '廠別代號'
                                    ,TO022 AS '交貨庫別'
                                    ,TO031 AS '備註'
                                    ,MA002 AS '廠商'
                                    ,MA003 AS '廠商全名'
                                    ,MA008 AS '廠商電話'
                                    ,MA013 AS '聯絡人'
                                    ,MA055 AS '付款條件'
                                    ,MA025 AS '付款'
                                    --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅  &&880210 &&88-11-25 OLD:預留C10
                                    ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '課稅別'
                                    ,MA047 AS '採購人員'
                                    ,MA010 AS '廠商傳真'
                                    ,MV002 AS '採購人'
                                    ,(TO024*TO017) AS '採購金額'

                                    ,TO110 AS '舊單位'
                                    ,TO109 AS '舊品號'
                                    ,TO135 AS '舊品名'
                                    ,TO136 AS '舊規格'
                                    ,TO124 AS '舊採購單價'
                                    ,TO117 AS '舊採購數量'
                                    ,TO145 AS '舊交易幣別'
                                    ,TO146 AS '舊匯率'
                                    ,TO121 AS '舊廠別代號'
                                    ,TO122 AS '舊交貨庫別'
                                    ,TO131 AS '舊備註'
                                    ,(TO124*TO117) AS '舊採購金額'
                                    ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '舊課稅別'

                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO024*TO017)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TO024*TO017)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '稅額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO024*TO017)) WHEN MA044=2 THEN CONVERT(INT,(TO024*TO017)+(TO024*TO017)*0.05) WHEN MA044=3 THEN (TO024*TO017) WHEN MA044=4 THEN (TO024*TO017) WHEN MA044=9 THEN (TO024*TO017) END )  AS '金額合計'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO124*TO117)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TO124*TO117)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '舊稅額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO124*TO117)) WHEN MA044=2 THEN CONVERT(INT,(TO124*TO117)+(TO124*TO117)*0.05) WHEN MA044=3 THEN (TO124*TO117) WHEN MA044=4 THEN (TO124*TO117) WHEN MA044=9 THEN (TO124*TO117) END )  AS '舊金額合計'
                                    ,TO005 AS '變更原因'
                                    ,TO113 AS '舊預交日期'
                                    ,(SELECT SUM(TA017) FROM [TK].dbo.MOCTA WHERE TA001=TO001 AND TA002=TO002)  AS '已交數量'

                                    FROM [TK].dbo.MOCTO
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TO033
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TO057
                                    WHERE TO001='{0}'
                                    AND TO004>='{1}' AND TO004<='{2}'
                                    ", TA001, SDAYS, EDAYS);



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView2.DataSource = ds1.Tables["ds1"];
                        dataGridView2.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                        // 設置欄位順序
                        dataGridView2.Columns["單別"].DisplayIndex = 0;
                        dataGridView2.Columns["單號"].DisplayIndex = 1;
                        dataGridView2.Columns["變更版次"].DisplayIndex = 2;
                        dataGridView2.Columns["廠商"].DisplayIndex = 3;
                        dataGridView2.Columns["品名"].DisplayIndex = 4;
                        dataGridView2.Columns["採購數量"].DisplayIndex = 5;


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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox9.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox10.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox15.Text = row.Cells["變更版次"].Value.ToString().Trim();
                }
            }
        }

        public void SETFASTREPORT2(string TO001, string TO002, string TO003, string P1, string P2, string P3, string P4)
        {
            report1 = new Report();
            string SQL = "";

            report1.Load(@"REPORT\託外採購變更單.frx");

            SQL = SETFASETSQL2(TO001, TO002,TO003);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;

            report1.SetParameterValue("公司電話", P1);
            report1.SetParameterValue("製表日期", DateTime.Now.ToString("yyyy/MM/dd"));
            report1.SetParameterValue("公司傳真", P2);
            report1.SetParameterValue("送貨地址", P3);
            report1.SetParameterValue("營業稅率", P4);

            Table.SelectCommand = SQL;

            report1.Preview = previewControl2;
            report1.Show();

        }
        public string SETFASETSQL2(string TO001, string TO002, string TO003)
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@" 
                                SELECT 
                                TO001 AS '單別'
                                ,TO002 AS '單號'
                                ,TO003 AS '變更版次'
                                ,CONVERT(NVARCHAR,CONVERT(datetime,TO004),111) AS '單據日期'
                                ,CONVERT(NVARCHAR,CONVERT(datetime,TO013),111) AS '到貨日'
                                ,CONVERT(NVARCHAR,CONVERT(datetime,TO012),111) AS '採購日期'
                                ,TO033 AS '廠商代號'
                                ,TO010 AS '單位'
                                ,TO009 AS '品號'
                                ,TO035 AS '品名'
                                ,TO036 AS '規格'
                                ,TO024 AS '採購單價'
                                ,TO017 AS '採購數量'
                                ,TO045 AS '交易幣別'
                                ,TO046 AS '匯率'
                                ,TO021 AS '廠別代號'
                                ,TO022 AS '交貨庫別'
                                ,TO031 AS '備註'
                                ,MA002 AS '廠商'
                                ,MA003 AS '廠商全名'
                                ,MA008 AS '廠商電話'
                                ,MA013 AS '聯絡人'
                                ,MA055 AS '付款條件'
                                ,MA025 AS '付款'
                                --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅  &&880210 &&88-11-25 OLD:預留C10
                                ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '課稅別'
                                ,MA047 AS '採購人員'
                                ,MA010 AS '廠商傳真'
                                ,MV002 AS '採購人'
                                ,(TO024*TO017) AS '採購金額'

                                ,TO110 AS '舊單位'
                                ,TO109 AS '舊品號'
                                ,TO135 AS '舊品名'
                                ,TO136 AS '舊規格'
                                ,TO124 AS '舊採購單價'
                                ,TO117 AS '舊採購數量'
                                ,TO145 AS '舊交易幣別'
                                ,TO146 AS '舊匯率'
                                ,TO121 AS '舊廠別代號'
                                ,TO122 AS '舊交貨庫別'
                                ,TO131 AS '舊備註'
                                ,(TO124*TO117) AS '舊採購金額'
                                ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '舊課稅別'

                                ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO024*TO017)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TO024*TO017)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '稅額'
                                ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO024*TO017)) WHEN MA044=2 THEN CONVERT(INT,(TO024*TO017)+(TO024*TO017)*0.05) WHEN MA044=3 THEN (TO024*TO017) WHEN MA044=4 THEN (TO024*TO017) WHEN MA044=9 THEN (TO024*TO017) END )  AS '金額合計'
                                ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO124*TO117)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TO124*TO117)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '舊稅額'
                                ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO124*TO117)) WHEN MA044=2 THEN CONVERT(INT,(TO124*TO117)+(TO124*TO117)*0.05) WHEN MA044=3 THEN (TO124*TO117) WHEN MA044=4 THEN (TO124*TO117) WHEN MA044=9 THEN (TO124*TO117) END )  AS '舊金額合計'
                                ,TO005 AS '變更原因'
                                ,TO113 AS '舊預交日期'
                                ,(SELECT SUM(TA017) FROM [TK].dbo.MOCTA WHERE TA001=TO001 AND TA002=TO002)  AS '已交數量'

                                FROM [TK].dbo.MOCTO
                                LEFT JOIN [TK].dbo.PURMA ON MA001=TO033
                                LEFT JOIN [TK].dbo.CMSMV ON MV001=TO057
                                WHERE TO001='{0}'
                                AND TO002='{1}'
                                AND TO003='{2}'

                                ", TO001, TO002, TO003);

            return FASTSQL.ToString();
        }

        public void SEARCH_MOCTC(string TA001, string SDAYS, string EDAYS)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
                                    TA001 AS '單別'
                                    ,TA002 AS '單號'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TA003),111) AS '單據日期'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TA010),111) AS '到貨日'
                                    ,TA032 AS '廠商代號'
                                    ,TA023 AS '單位'
                                    ,TA006 AS '品號'
                                    ,TA034 AS '品名'
                                    ,TA035 AS '規格'
                                    ,TA022 AS '採購單價'
                                    ,TA015 AS '採購數量'
                                    ,TA042 AS '交易幣別'
                                    ,TA043 AS '匯率'
                                    ,TA019 AS '廠別代號'
                                    ,TA020 AS '交貨庫別'
                                    ,TA029 AS '備註'
                                    ,MA002 AS '廠商'
                                    ,MA003 AS '廠商全名'
                                    ,MA008 AS '廠商電話'
                                    ,MA013 AS '聯絡人'
                                    ,MA055 AS '付款條件'
                                    ,MA025 AS '付款'
                                    --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅  &&880210 &&88-11-25 OLD:預留C10
                                    ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '課稅別'
                                    ,MA047 AS '採購人員'
                                    ,MA010 AS '廠商傳真'
                                    ,MV002 AS '採購人'
                                    ,(TA022*TA015) AS '採購金額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '稅額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)+(TA022*TA015)*0.05) WHEN MA044=3 THEN (TA022*TA015) WHEN MA044=4 THEN (TA022*TA015) WHEN MA044=9 THEN (TA022*TA015) END )  AS '金額合計'
                                    ,TA003

                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TA032
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=MA047
                                    WHERE TA001='{0}'
                                    AND TA003>='{1}' AND TA003<='{2}'
                                    ", TA001, SDAYS, EDAYS);



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();

                
                dataGridView3.DataSource = null;
                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView3.DataSource = ds1.Tables["ds1"];
                    dataGridView3.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                    // 設置欄位順序
                    dataGridView3.Columns["單別"].DisplayIndex = 0;
                    dataGridView3.Columns["單號"].DisplayIndex = 1;
                    dataGridView3.Columns["廠商"].DisplayIndex = 2;
                    dataGridView3.Columns["品名"].DisplayIndex = 3;
                    dataGridView3.Columns["採購數量"].DisplayIndex = 4;


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
            string TC045;
            textBox17.Text = null;
            textBox18.Text = null;
            textBox19.Text = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox17.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox18.Text = row.Cells["單號"].Value.ToString().Trim();
                    //textBox19.Text = row.Cells["TA003"].Value.ToString().Trim();

                    TC045 = textBox17.Text.Trim() + textBox18.Text.Trim();
                    //是否已產生託外採購單
                    SERACH_PURTC(TC045);
                }

            }
        }
        public string GETMAXTC002(string TC001, string TC003)
        {
            string TC002;
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                               SELECT ISNULL(MAX(TC002),'00000000000') AS TC002
                                FROM [TK].[dbo].[PURTC]
                                WHERE  TC001='{0}' AND TC002 LIKE '%{1}%'  
                                        ", TC001, TC003);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        TC002 = SETTC002(ds1.Tables["ds1"].Rows[0]["TC002"].ToString(),TC003);

                        return TC002;

                    }
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

        public string SETTC002(string TC002,string TC003)
        {
            if (TC002.Equals("00000000000"))
            {
                return TC003 + "001";
            }

            else
            {
                int serno = Convert.ToInt16(TC002.Substring(8, 3));
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(3, '0');
                return TC003 + temp.ToString();
            }
        }

        public string GETMAXTE003(string TE001,string TE002)
        {
            string TE003;
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();


                sbSql.AppendFormat(@" 
                                    SELECT ISNULL(MAX(TE003),'0000') AS TE003
                                    FROM [TK].[dbo].[PURTE]
                                    WHERE  TE001='{0}' AND TE002 LIKE '%{1}%'  
                                        ", TE001, TE002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                if (ds1.Tables["ds1"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds1.Tables["ds1"].Rows.Count >= 1)
                    {
                        TE003 = SETTE003(ds1.Tables["ds1"].Rows[0]["TE003"].ToString());

                        return TE003;

                    }
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
        public string SETTE003(string TE003)
        {
            if (TE003.Equals("0000"))
            {
                return "0001";
            }

            else
            {
                int serno = Convert.ToInt16(TE003);
                serno = serno + 1;
                string temp = serno.ToString();
                temp = temp.PadLeft(4, '0');
                return  temp.ToString();
            }
        }


        public void ADD_PURTC_PURTD(string TA001,string TA002,string TC001,string TC002,string TC003,string TC024)
        {
            DATA_SET ERPDATA = new DATA_SET();
            ERPDATA.COMPANY = "TK";
            ERPDATA.CREATOR = "070002";
            ERPDATA.USR_GROUP = "112000";
            ERPDATA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            ERPDATA.MODIFIER = "070002";
            ERPDATA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            ERPDATA.FLAG = "0";
            ERPDATA.CREATE_TIME = DateTime.Now.ToString("HH:mm:ss");
            ERPDATA.MODI_TIME = DateTime.Now.ToString("HH:mm:ss");
            ERPDATA.TRANS_TYPE = "P001";
            ERPDATA.TRANS_NAME = "PURMI07";
            ERPDATA.sync_date = "";
            ERPDATA.sync_time = ""; 
            ERPDATA.sync_mark = ""; 
            ERPDATA.sync_count = "0"; 
            ERPDATA.DataUser = "";
            ERPDATA.DataGroup = "112000";

            //將來源的託外製令單號，放在託外採購單的TC045
            //合約編號
            string TC045 = TA001.Trim() + TA002.Trim();

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

                                    INSERT INTO [TK].dbo.PURTC
                                    (
                                    COMPANY
                                    ,CREATOR
                                    ,USR_GROUP
                                    ,CREATE_DATE
                                    ,MODIFIER
                                    ,MODI_DATE
                                    ,FLAG
                                    ,CREATE_TIME
                                    ,MODI_TIME
                                    ,TRANS_TYPE
                                    ,TRANS_NAME
                                    ,sync_date
                                    ,sync_time
                                    ,sync_mark
                                    ,sync_count
                                    ,DataUser
                                    ,DataGroup
                                    ,TC001
                                    ,TC002
                                    ,TC003
                                    ,TC004
                                    ,TC005
                                    ,TC006
                                    ,TC007
                                    ,TC008
                                    ,TC009
                                    ,TC010
                                    ,TC011
                                    ,TC012
                                    ,TC013
                                    ,TC014
                                    ,TC015
                                    ,TC016
                                    ,TC017
                                    ,TC018
                                    ,TC019
                                    ,TC020
                                    ,TC021
                                    ,TC022
                                    ,TC023
                                    ,TC024
                                    ,TC025
                                    ,TC026
                                    ,TC027
                                    ,TC028
                                    ,TC029
                                    ,TC030
                                    ,TC031
                                    ,TC032
                                    ,TC033
                                    ,TC034
                                    ,TC035
                                    ,TC036
                                    ,TC037
                                    ,TC038
                                    ,TC039
                                    ,TC040
                                    ,TC041
                                    ,TC042
                                    ,TC043
                                    ,TC044
                                    ,TC045
                                    ,TC046
                                    ,TC047
                                    ,TC048
                                    ,TC049
                                    ,TC050
                                    ,TC051
                                    ,TC052
                                    ,TC053
                                    ,TC054
                                    ,TC055
                                    ,TC056
                                    ,TC057
                                    ,TC058
                                    ,TC059
                                    ,TC060
                                    ,TC061
                                    ,TC062
                                    ,TC063
                                    ,TC064
                                    ,TC065
                                    ,TC066
                                    ,TC067
                                    ,TC068
                                    ,TC069
                                    ,TC070
                                    ,TC071
                                    ,TC072
                                    ,TC073
                                    ,TC074
                                    ,TC075
                                    ,TC076
                                    ,TC077
                                    ,TC078
                                    ,TC079
                                    ,TC080
                                    ,UDF01
                                    ,UDF02
                                    ,UDF03
                                    ,UDF04
                                    ,UDF05
                                    ,UDF06
                                    ,UDF07
                                    ,UDF08
                                    ,UDF09
                                    ,UDF10
                                    )
                                    SELECT 
                                    '{0}' COMPANY
                                    ,'{1}' CREATOR
                                    ,'{2}' USR_GROUP
                                    ,'{3}' CREATE_DATE
                                    ,'{4}' MODIFIER
                                    ,'{5}' MODI_DATE
                                    ,'{6}' FLAG
                                    ,'{7}' CREATE_TIME
                                    ,'{8}' MODI_TIME
                                    ,'{9}' TRANS_TYPE
                                    ,'{10}' TRANS_NAME
                                    ,'{11}' sync_date
                                    ,'{12}' sync_time
                                    ,'{13}' sync_mark
                                    ,'{14}' sync_count
                                    ,'{15}' DataUser
                                    ,'{16}' DataGroup
                                    ,'{17}' TC001
                                    ,'{18}' TC002
                                    ,{22} TC003
                                    ,TA032 TC004
                                    ,TA042 TC005
                                    ,TA043 TC006
                                    ,'' TC007
                                    ,MA025 TC008
                                    ,TA029 TC009
                                    ,TA019 TC010
                                    ,MA047 TC011
                                    ,'1' TC012
                                    ,'1' TC013
                                    ,'N' TC014
                                    ,'' TC015
                                    ,'' TC016
                                    ,'' TC017
                                    ,MA044 TC018
                                    ,(TA022*TA015) TC019
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TA022*TA015)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TA022*TA015)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  TC020
                                    ,'嘉義縣大林鎮大埔美園區五路3號' TC021
                                    ,'' TC022
                                    ,TA015 TC023
                                    ,{22} TC024
                                    ,'' TC025
                                    ,'0.0500' TC026
                                    ,MA055 TC027
                                    ,'0' TC028
                                    ,'0' TC029
                                    ,'N' TC030
                                    ,'0' TC031
                                    ,'' TC032
                                    ,'N' TC033
                                    ,'' TC034
                                    ,'' TC035
                                    ,'' TC036
                                    ,'05-2956520' TC037
                                    ,'05-2956519' TC038
                                    ,'' TC039
                                    ,'' TC040
                                    ,'' TC041
                                    ,'0' TC042
                                    ,'0' TC043
                                    ,'' TC044
                                    ,'{21}' TC045
                                    ,'' TC046
                                    ,'' TC047
                                    ,'' TC048
                                    ,'' TC049
                                    ,'N' TC050
                                    ,'' TC051
                                    ,'' TC052
                                    ,'' TC053
                                    ,'' TC054
                                    ,'' TC055
                                    ,'' TC056
                                    ,'' TC057
                                    ,'' TC058
                                    ,'' TC059
                                    ,'' TC060
                                    ,'' TC061
                                    ,'' TC062
                                    ,'' TC063
                                    ,'' TC064
                                    ,'' TC065
                                    ,'0' TC066
                                    ,'' TC067
                                    ,'' TC068
                                    ,'' TC069
                                    ,'0' TC070
                                    ,'0' TC071
                                    ,'0' TC072
                                    ,'0' TC073
                                    ,'0' TC074
                                    ,'' TC075
                                    ,'' TC076
                                    ,'' TC077
                                    ,'' TC078
                                    ,'' TC079
                                    ,'' TC080
                                    ,'' UDF01
                                    ,'' UDF02
                                    ,'' UDF03
                                    ,'' UDF04
                                    ,'' UDF05
                                    ,'0' UDF06
                                    ,'0' UDF07
                                    ,'0' UDF08
                                    ,'0' UDF09
                                    ,'0'  UDF10
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TA032
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=MA047
                                    WHERE TA001='{19}'
                                    AND TA002='{20}'
                                        ", ERPDATA.COMPANY 
                                        ,ERPDATA.CREATOR 
                                        ,ERPDATA.USR_GROUP 
                                        ,ERPDATA.CREATE_DATE 
                                        ,ERPDATA.MODIFIER 
                                        ,ERPDATA.MODI_DATE
                                        ,ERPDATA.FLAG 
                                        ,ERPDATA.CREATE_TIME
                                        ,ERPDATA.MODI_TIME
                                        ,ERPDATA.TRANS_TYPE 
                                        ,ERPDATA.TRANS_NAME 
                                        ,ERPDATA.sync_date 
                                        ,ERPDATA.sync_time
                                        ,ERPDATA.sync_mark
                                        ,ERPDATA.sync_count
                                        ,ERPDATA.DataUser
                                        ,ERPDATA.DataGroup 
                                        ,TC001
                                        ,TC002
                                        ,TA001
                                        ,TA002
                                        ,TC045
                                        ,TC024
                                        );

                sbSql.AppendFormat(@" 

                                    INSERT INTO [TK].[dbo].[PURTD]
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
                                    ,[TD001]
                                    ,[TD002]
                                    ,[TD003]
                                    ,[TD004]
                                    ,[TD005]
                                    ,[TD006]
                                    ,[TD007]
                                    ,[TD008]
                                    ,[TD009]
                                    ,[TD010]
                                    ,[TD011]
                                    ,[TD012]
                                    ,[TD013]
                                    ,[TD014]
                                    ,[TD015]
                                    ,[TD016]
                                    ,[TD017]
                                    ,[TD018]
                                    ,[TD019]
                                    ,[TD020]
                                    ,[TD021]
                                    ,[TD022]
                                    ,[TD023]
                                    ,[TD024]
                                    ,[TD025]
                                    ,[TD026]
                                    ,[TD027]
                                    ,[TD028]
                                    ,[TD029]
                                    ,[TD030]
                                    ,[TD031]
                                    ,[TD032]
                                    ,[TD033]
                                    ,[TD034]
                                    ,[TD035]
                                    ,[TD036]
                                    ,[TD037]
                                    ,[TD038]
                                    ,[TD039]
                                    ,[TD040]
                                    ,[TD041]
                                    ,[TD042]
                                    ,[TD043]
                                    ,[TD044]
                                    ,[TD045]
                                    ,[TD046]
                                    ,[TD047]
                                    ,[TD048]
                                    ,[TD049]
                                    ,[TD050]
                                    ,[TD051]
                                    ,[TD052]
                                    ,[TD053]
                                    ,[TD054]
                                    ,[TD055]
                                    ,[TD056]
                                    ,[TD057]
                                    ,[TD058]
                                    ,[TD059]
                                    ,[TD060]
                                    ,[TD061]
                                    ,[TD062]
                                    ,[TD063]
                                    ,[TD064]
                                    ,[TD065]
                                    ,[TD066]
                                    ,[TD067]
                                    ,[TD068]
                                    ,[TD069]
                                    ,[TD070]
                                    ,[TD071]
                                    ,[TD072]
                                    ,[TD073]
                                    ,[TD074]
                                    ,[TD075]
                                    ,[TD076]
                                    ,[TD077]
                                    ,[TD078]
                                    ,[TD079]
                                    ,[TD080]
                                    ,[TD081]
                                    ,[TD082]
                                    ,[TD083]
                                    ,[TD084]
                                    ,[TD085]
                                    ,[TD086]
                                    ,[TD087]
                                    ,[TD088]
                                    ,[TD089]
                                    ,[TD090]
                                    ,[TD091]
                                    ,[TD092]
                                    ,[TD093]
                                    ,[TD094]
                                    ,[TD095]
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
                                    '{0}' COMPANY
                                    ,'{1}' CREATOR
                                    ,'{2}' USR_GROUP
                                    ,'{3}' CREATE_DATE
                                    ,'{4}' MODIFIER
                                    ,'{5}' MODI_DATE
                                    ,'{6}' FLAG
                                    ,'{7}' CREATE_TIME
                                    ,'{8}' MODI_TIME
                                    ,'{9}' TRANS_TYPE
                                    ,'{10}' TRANS_NAME
                                    ,'{11}' sync_date
                                    ,'{12}' sync_time
                                    ,'{13}' sync_mark
                                    ,'{14}' sync_count
                                    ,'{15}' DataUser
                                    ,'{16}' DataGroup
                                    ,'{17}' TD001
                                    ,'{18}' TD002
                                    ,'0001 'TD003
                                    ,TA006 TD004
                                    ,TA034 TD005
                                    ,TA035 TD006
                                    ,TA020 TD007
                                    ,TA015 TD008
                                    ,TA023 TD009
                                    ,TA022 TD010
                                    ,(TA022*TA015)  TD011
                                    ,TA010 TD012
                                    ,'' TD013
                                    ,TA029 TD014
                                    ,0 TD015
                                    ,'N' TD016
                                    ,'' TD017
                                    ,'N' TD018
                                    ,'0' TD019
                                    ,'' TD020
                                    ,'' TD021
                                    ,'' TD022
                                    ,'' TD023
                                    ,'' TD024
                                    ,'N' TD025
                                    ,'' TD026
                                    ,'' TD027
                                    ,'' TD028
                                    ,'' TD029
                                    ,'0' TD030
                                    ,'0' TD031
                                    ,'' TD032
                                    ,'' TD033
                                    ,'0' TD034
                                    ,'0' TD035
                                    ,'' TD036
                                    ,'' TD037
                                    ,'' TD038
                                    ,'' TD039
                                    ,'' TD040
                                    ,'' TD041
                                    ,'' TD042
                                    ,'' TD043
                                    ,'' TD044
                                    ,'' TD045
                                    ,'' TD046
                                    ,'' TD047
                                    ,'0' TD048
                                    ,'0' TD049
                                    ,'' TD050
                                    ,'' TD051
                                    ,'' TD052
                                    ,'' TD053
                                    ,'' TD054
                                    ,'' TD055
                                    ,'' TD056
                                    ,'0' TD057
                                    ,'9' TD058
                                    ,'' TD059
                                    ,'' TD060
                                    ,'' TD061
                                    ,'2' TD062
                                    ,'' TD063
                                    ,'' TD064
                                    ,'' TD065
                                    ,'' TD066
                                    ,'' TD067
                                    ,'' TD068
                                    ,'0' TD069
                                    ,'' TD070
                                    ,'' TD071
                                    ,'N' TD072
                                    ,'' TD073
                                    ,'' TD074
                                    ,'0' TD075
                                    ,'' TD076
                                    ,'' TD077
                                    ,'' TD078
                                    ,'0' TD079
                                    ,'0' TD080
                                    ,'' TD081
                                    ,'0' TD082
                                    ,'' TD083
                                    ,'0' TD084
                                    ,'1' TD085
                                    ,'0' TD086
                                    ,'0' TD087
                                    ,'0' TD088
                                    ,'0' TD089
                                    ,'' TD090
                                    ,'' TD091
                                    ,'' TD092
                                    ,'' TD093
                                    ,'' TD094
                                    ,'' TD095
                                    ,'' UDF01
                                    ,'' UDF02
                                    ,'' UDF03
                                    ,'' UDF04
                                    ,'' UDF05
                                    ,'0' UDF06
                                    ,'0' UDF07
                                    ,'0' UDF08
                                    ,'0' UDF09
                                    ,'0' UDF10
                                    FROM [TK].dbo.MOCTA
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TA032
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=MA047
                                    WHERE TA001='{19}'
                                    AND TA002='{20}'
                                    ", ERPDATA.COMPANY
                                        , ERPDATA.CREATOR
                                        , ERPDATA.USR_GROUP
                                        , ERPDATA.CREATE_DATE
                                        , ERPDATA.MODIFIER
                                        , ERPDATA.MODI_DATE
                                        , ERPDATA.FLAG
                                        , ERPDATA.CREATE_TIME
                                        , ERPDATA.MODI_TIME
                                        , ERPDATA.TRANS_TYPE
                                        , ERPDATA.TRANS_NAME
                                        , ERPDATA.sync_date
                                        , ERPDATA.sync_time
                                        , ERPDATA.sync_mark
                                        , ERPDATA.sync_count
                                        , ERPDATA.DataUser
                                        , ERPDATA.DataGroup
                                        , TC001
                                        , TC002
                                        , TA001
                                        , TA002
                                    );


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

        public void SERACH_PURTC(string TC045)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
                                    TC001  AS '採購單別'
                                    ,TC002 AS  '採購單號'
                                    FROM  [TK].dbo.PURTC
                                    WHERE TC045='{0}'
                                    ", TC045);



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                dataGridView4.DataSource = null;
                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView4.DataSource = ds1.Tables["ds1"];
                    dataGridView4.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];  


                }
            }
            catch
            {

            }
            finally
            {

            }
        }
        public void SERACH_PURTE(string TE015)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
                                    TE001  AS '採購變更單別'
                                    ,TE002 AS  '採購變更單號'
                                    ,TE003 AS  '版次'
                                    FROM  [TK].dbo.PURTE
                                    WHERE TE015='{0}'
                                    ", TE015);



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                dataGridView6.DataSource = null;
                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView6.DataSource = ds1.Tables["ds1"];
                    dataGridView6.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];  


                }
            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCH_MOCTO_V2(string TA001,string SDAYS,string EDAYS)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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
                                    TO001 AS '單別'
                                    ,TO002 AS '單號'
                                    ,TO003 AS '變更版次'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TO004),111) AS '單據日期'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TO013),111) AS '到貨日'
                                    ,CONVERT(NVARCHAR,CONVERT(datetime,TO012),111) AS '採購日期'
                                    ,TO033 AS '廠商代號'
                                    ,TO010 AS '單位'
                                    ,TO009 AS '品號'
                                    ,TO035 AS '品名'
                                    ,TO036 AS '規格'
                                    ,TO024 AS '採購單價'
                                    ,TO017 AS '採購數量'
                                    ,TO045 AS '交易幣別'
                                    ,TO046 AS '匯率'
                                    ,TO021 AS '廠別代號'
                                    ,TO022 AS '交貨庫別'
                                    ,TO031 AS '備註'
                                    ,MA002 AS '廠商'
                                    ,MA003 AS '廠商全名'
                                    ,MA008 AS '廠商電話'
                                    ,MA013 AS '聯絡人'
                                    ,MA055 AS '付款條件'
                                    ,MA025 AS '付款'
                                    --1.應稅內含、2.應稅外加、3.零稅率、4.免稅、9.不計稅  &&880210 &&88-11-25 OLD:預留C10
                                    ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '課稅別'
                                    ,MA047 AS '採購人員'
                                    ,MA010 AS '廠商傳真'
                                    ,MV002 AS '採購人'
                                    ,(TO024*TO017) AS '採購金額'

                                    ,TO110 AS '舊單位'
                                    ,TO109 AS '舊品號'
                                    ,TO135 AS '舊品名'
                                    ,TO136 AS '舊規格'
                                    ,TO124 AS '舊採購單價'
                                    ,TO117 AS '舊採購數量'
                                    ,TO145 AS '舊交易幣別'
                                    ,TO146 AS '舊匯率'
                                    ,TO121 AS '舊廠別代號'
                                    ,TO122 AS '舊交貨庫別'
                                    ,TO131 AS '舊備註'
                                    ,(TO124*TO117) AS '舊採購金額'
                                    ,(CASE WHEN MA044=1 THEN '應稅內含' WHEN MA044=2 THEN '應稅外加' WHEN MA044=3 THEN '零稅率' WHEN MA044=4 THEN '免稅' WHEN MA044=9 THEN '不計稅'  END )  AS '舊課稅別'

                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO024*TO017)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TO024*TO017)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '稅額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO024*TO017)) WHEN MA044=2 THEN CONVERT(INT,(TO024*TO017)+(TO024*TO017)*0.05) WHEN MA044=3 THEN (TO024*TO017) WHEN MA044=4 THEN (TO024*TO017) WHEN MA044=9 THEN (TO024*TO017) END )  AS '金額合計'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO124*TO117)/1.05) WHEN MA044=2 THEN CONVERT(INT,(TO124*TO117)*0.05) WHEN MA044=3 THEN 0 WHEN MA044=4 THEN 0 WHEN MA044=9 THEN 0 END )  AS '舊稅額'
                                    ,(CASE WHEN MA044=1 THEN CONVERT(INT,(TO124*TO117)) WHEN MA044=2 THEN CONVERT(INT,(TO124*TO117)+(TO124*TO117)*0.05) WHEN MA044=3 THEN (TO124*TO117) WHEN MA044=4 THEN (TO124*TO117) WHEN MA044=9 THEN (TO124*TO117) END )  AS '舊金額合計'
                                    ,TO005 AS '變更原因'
                                    ,TO113 AS '舊預交日期'
                                    ,(SELECT SUM(TA017) FROM [TK].dbo.MOCTA WHERE TA001=TO001 AND TA002=TO002)  AS '已交數量'
                                    ,(SELECT TOP 1 TC001 FROM [TK].dbo.PURTC WHERE TC045=REPLACE(TO001+TO002,' ','') ORDER BY TC003 DESC ) AS TC001
                                    ,(SELECT TOP 1 TC002 FROM [TK].dbo.PURTC WHERE TC045=REPLACE(TO001+TO002,' ','') ORDER BY TC003 DESC ) AS TC002

                                    FROM [TK].dbo.MOCTO
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TO033
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TO057
                                    WHERE TO001='{0}'
                                    AND TO004>='{1}' AND TO004<='{2}'
                                    ", TA001, SDAYS, EDAYS);



                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "ds1");
                sqlConn.Close();


                dataGridView5.DataSource = null;
                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView5.DataSource = ds1.Tables["ds1"];
                    dataGridView5.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];
                    // 設置欄位順序
                    dataGridView5.Columns["單別"].DisplayIndex = 0;
                    dataGridView5.Columns["單號"].DisplayIndex = 1;
                    dataGridView5.Columns["變更版次"].DisplayIndex = 2;
                    dataGridView5.Columns["廠商"].DisplayIndex = 3;
                    dataGridView5.Columns["品名"].DisplayIndex = 4;
                    dataGridView5.Columns["採購數量"].DisplayIndex = 5;


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
            textBox21.Text = null;
            textBox22.Text = null;
            textBox23.Text = null;
            textBox24.Text = null;
            textBox25.Text = null;

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    textBox21.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox22.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox23.Text = row.Cells["變更版次"].Value.ToString().Trim();
                    textBox24.Text = row.Cells["TC001"].Value.ToString().Trim();
                    textBox25.Text = row.Cells["TC002"].Value.ToString().Trim();

                    string TE015 = textBox21.Text.Trim() + textBox22.Text.Trim() + textBox23.Text.Trim();
                    //是否已產生託外採購單
                    SERACH_PURTE(TE015);
                }

            }
        }

        public void ADD_PURTE_PURTF(string TE001,string TE002,string TE003,string TO001,string TO002,string TO003)
        {

            DATA_SET ERPDATA = new DATA_SET();
            ERPDATA.COMPANY = "TK";
            ERPDATA.CREATOR = "070002";
            ERPDATA.USR_GROUP = "112000";
            ERPDATA.CREATE_DATE = DateTime.Now.ToString("yyyyMMdd");
            ERPDATA.MODIFIER = "070002";
            ERPDATA.MODI_DATE = DateTime.Now.ToString("yyyyMMdd");
            ERPDATA.FLAG = "0";
            ERPDATA.CREATE_TIME = DateTime.Now.ToString("HH:mm:ss");
            ERPDATA.MODI_TIME = DateTime.Now.ToString("HH:mm:ss");
            ERPDATA.TRANS_TYPE = "P001";
            ERPDATA.TRANS_NAME = "PURMI07";
            ERPDATA.sync_date = "";
            ERPDATA.sync_time = "";
            ERPDATA.sync_mark = "";
            ERPDATA.sync_count = "0";
            ERPDATA.DataUser = "";
            ERPDATA.DataGroup = "112000";

            //將來源的託外製令單號，放在託外採購單的TC045
            //合約編號
            string TE015 = TO001.Trim() + TO002.Trim() + TO003.Trim();
            string TE004 = TE002.Substring(0,8);
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
                                    '{0}' COMPANY
                                    ,'{1}' CREATOR
                                    ,'{2}' USR_GROUP
                                    ,'{3}' CREATE_DATE
                                    ,'{4}' MODIFIER
                                    ,'{5}' MODI_DATE
                                    ,'{6}' FLAG
                                    ,'{7}' CREATE_TIME
                                    ,'{8}' MODI_TIME
                                    ,'{9}' TRANS_TYPE
                                    ,'{10}' TRANS_NAME
                                    ,'{11}' sync_date
                                    ,'{12}' sync_time
                                    ,'{13}' sync_mark
                                    ,'{14}' sync_count
                                    ,'{15}' DataUser
                                    ,'{16}' DataGroup
                                    ,'{17}' TE001
                                    ,'{18}' TE002
                                    ,'{19}' TE003
                                    ,'{20}' TE004
                                    ,TO033 TE005
                                    ,TO031 TE006
                                    ,TO045 TE007
                                    ,TO046 TE008
                                    ,'' TE009
                                    ,MA025 TE010
                                    ,TO004 TE011
                                    ,'N' TE012
                                    ,'' TE013
                                    ,'' TE014
                                    ,'{24}' TE015
                                    ,'0' TE016
                                    ,'N' TE017
                                    ,MA044 TE018
                                    ,'嘉義縣大林鎮大埔美園區五路3號' TE019
                                    ,'' TE020
                                    ,'' TE021
                                    ,'0.0500' TE022
                                    ,MA055 TE023
                                    ,'0' TE024
                                    ,'N' TE025
                                    ,'0' TE026
                                    ,'' TE027
                                    ,'N' TE028
                                    ,'' TE029
                                    ,'' TE030
                                    ,'' TE031
                                    ,'' TE032
                                    ,'' TE033
                                    ,'0' TE034
                                    ,'0' TE035
                                    ,'' TE036
                                    ,'070002' TE037
                                    ,'' TE038
                                    ,'' TE039
                                    ,'' TE040
                                    ,'N' TE041
                                    ,'' TE042
                                    ,'' TE043
                                    ,'' TE045
                                    ,'' TE046
                                    ,'' TE047
                                    ,'' TE048
                                    ,'' TE103
                                    ,TO045 TE107
                                    ,TO046 TE108
                                    ,'' TE109
                                    ,MA025 TE110
                                    ,'' TE113
                                    ,'' TE114
                                    ,'' TE115
                                    ,MA044 TE118
                                    ,'嘉義縣大林鎮大埔美園區五路3號' TE119
                                    ,'' TE120
                                    ,'0.0500' TE121
                                    ,MA055 TE122
                                    ,'0' TE123
                                    ,'' TE124
                                    ,'' TE125
                                    ,'0' TE134
                                    ,'0' TE135
                                    ,'' TE136
                                    ,'' TE137
                                    ,'' TE138
                                    ,'' TE139
                                    ,'' TE140
                                    ,'N' TE141
                                    ,'' TE142
                                    ,'' TE143
                                    ,'' TE144
                                    ,'' TE145
                                    ,'' TE146
                                    ,'' TE147
                                    ,'070002' TE148
                                    ,'0' TE149
                                    ,'0' TE150
                                    ,'0' TE151
                                    ,'0' TE152
                                    ,'' TE153
                                    ,'' TE154
                                    ,'' TE155
                                    ,'' TE156
                                    ,'' TE157
                                    ,'' TE158
                                    ,'' TE159
                                    ,'' TE160
                                    ,'' TE161
                                    ,'' TE162
                                    ,'' UDF01
                                    ,'' UDF02
                                    ,'' UDF03
                                    ,'' UDF04
                                    ,'' UDF05
                                    ,'0' UDF06
                                    ,'0' UDF07
                                    ,'0' UDF08
                                    ,'0' UDF09
                                    ,'0' UDF10
                                    FROM [TK].dbo.MOCTO
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TO033
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TO057

                                    WHERE TO001='{21}'
                                    AND TO002='{22}'
                                    AND TO003='{23}'
                                    ", ERPDATA.COMPANY
                                        , ERPDATA.CREATOR
                                        , ERPDATA.USR_GROUP
                                        , ERPDATA.CREATE_DATE
                                        , ERPDATA.MODIFIER
                                        , ERPDATA.MODI_DATE
                                        , ERPDATA.FLAG
                                        , ERPDATA.CREATE_TIME
                                        , ERPDATA.MODI_TIME
                                        , ERPDATA.TRANS_TYPE
                                        , ERPDATA.TRANS_NAME
                                        , ERPDATA.sync_date
                                        , ERPDATA.sync_time
                                        , ERPDATA.sync_mark
                                        , ERPDATA.sync_count
                                        , ERPDATA.DataUser
                                        , ERPDATA.DataGroup
                                        , TE001
                                        , TE002
                                        , TE003
                                        , TE004
                                        , TO001
                                        , TO002
                                        , TO003
                                        , TE015
                                        );

                sbSql.AppendFormat(@"
                                    INSERT INTO  [TK].[dbo].[PURTF]
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
                                    '{0}' COMPANY
                                    ,'{1}' CREATOR
                                    ,'{2}' USR_GROUP
                                    ,'{3}' CREATE_DATE
                                    ,'{4}' MODIFIER
                                    ,'{5}' MODI_DATE
                                    ,'{6}' FLAG
                                    ,'{7}' CREATE_TIME
                                    ,'{8}' MODI_TIME
                                    ,'{9}' TRANS_TYPE
                                    ,'{10}' TRANS_NAME
                                    ,'{11}' sync_date
                                    ,'{12}' sync_time
                                    ,'{13}' sync_mark
                                    ,'{14}' sync_count
                                    ,'{15}' DataUser
                                    ,'{16}' DataGroup
                                    ,'{17}' TF001
                                    ,'{18}' TF002
                                    ,'{19}' TF003
                                    ,'0001' TF004
                                    ,TO009 TF005
                                    ,TO035 TF006
                                    ,TO036 TF007
                                    ,TO022 TF008
                                    ,TO017 TF009
                                    ,TO010 TF010
                                    ,TO024 TF011
                                    ,(TO024*TO017) TF012
                                    ,TO013 TF013
                                    ,'N' TF014
                                    ,'0' TF015
                                    ,'N' TF016
                                    ,TO031 TF017
                                    ,'0' TF018
                                    ,'' TF019
                                    ,'' TF020
                                    ,'N' TF021
                                    ,'' TF022
                                    ,'' TF023
                                    ,'0' TF024
                                    ,'' TF025
                                    ,'' TF026
                                    ,'' TF027
                                    ,'' TF028
                                    ,'' TF029
                                    ,'' TF030
                                    ,'' TF031
                                    ,'' TF032
                                    ,'' TF033
                                    ,'' TF034
                                    ,'' TF035
                                    ,'0' TF036
                                    ,'0' TF037
                                    ,'' TF038
                                    ,'' TF039
                                    ,'' TF040
                                    ,'0' TF041
                                    ,'0001' TF104
                                    ,TO109 TF105
                                    ,TO135 TF106
                                    ,TO136 TF107
                                    ,TO122 TF108
                                    ,TO117 TF109
                                    ,TO110 TF110
                                    ,TO124 TF111
                                    ,(TO124*TO117) TF112
                                    ,TO013 TF113
                                    ,'N' TF114
                                    ,'0' TF118
                                    ,'' TF119
                                    ,'' TF120
                                    ,'N' TF121
                                    ,'' TF122
                                    ,'' TF123
                                    ,'0' TF124
                                    ,'0' TF125
                                    ,'' TF126
                                    ,'' TF127
                                    ,'0' TF128
                                    ,'0' TF129
                                    ,'0' TF130
                                    ,'0' TF131
                                    ,'' TF132
                                    ,'' TF133
                                    ,'' TF134
                                    ,'' TF135
                                    ,'0' TF136
                                    ,'0' TF137
                                    ,'' TF138
                                    ,'' TF139
                                    ,'' TF140
                                    ,'0' TF141
                                    ,'' TF142
                                    ,'' TF143
                                    ,'' TF144
                                    ,'2' TF145
                                    ,'2' TF146
                                    ,'' TF147
                                    ,'' TF148
                                    ,'' TF149
                                    ,'' TF150
                                    ,'' TF151
                                    ,'0' TF152
                                    ,'' TF153
                                    ,'0' TF154
                                    ,'' TF155
                                    ,'0' TF156
                                    ,'' TF157
                                    ,'0' TF158
                                    ,'' TF159
                                    ,'0' TF160
                                    ,'1' TF161
                                    ,'0' TF162
                                    ,'1' TF163
                                    ,'0' TF164
                                    ,'0' TF165
                                    ,'0' TF166
                                    ,'0' TF167
                                    ,'' TF168
                                    ,'' TF169
                                    ,'' TF170
                                    ,'' TF171
                                    ,'' TF172
                                    ,'' TF173
                                    ,'' UDF01
                                    ,'' UDF02
                                    ,'' UDF03
                                    ,'' UDF04
                                    ,'' UDF05
                                    ,'0' UDF06
                                    ,'0' UDF07
                                    ,'0' UDF08
                                    ,'0' UDF09
                                    ,'0' UDF10
                                    FROM [TK].dbo.MOCTO
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TO033
                                    LEFT JOIN [TK].dbo.CMSMV ON MV001=TO057

                                    WHERE TO001='{20}'
                                    AND TO002='{21}' 
                                    AND TO003='{22}'

                                    ", ERPDATA.COMPANY
                                        , ERPDATA.CREATOR
                                        , ERPDATA.USR_GROUP
                                        , ERPDATA.CREATE_DATE
                                        , ERPDATA.MODIFIER
                                        , ERPDATA.MODI_DATE
                                        , ERPDATA.FLAG
                                        , ERPDATA.CREATE_TIME
                                        , ERPDATA.MODI_TIME
                                        , ERPDATA.TRANS_TYPE
                                        , ERPDATA.TRANS_NAME
                                        , ERPDATA.sync_date
                                        , ERPDATA.sync_time
                                        , ERPDATA.sync_mark
                                        , ERPDATA.sync_count
                                        , ERPDATA.DataUser
                                        , ERPDATA.DataGroup
                                        , TE001
                                        , TE002
                                        , TE003                                       
                                        , TO001
                                        , TO002
                                        , TO003
                                        );

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

        public void SETFASTREPORT(string SDAY,string EDAY,string MB001,string TA006)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDAY, EDAY, MB001, TA006);
            Report report1 = new Report();
            report1.Load(@"REPORT\託外採購-製令-進貨.frx");

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

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDAY, string EDAY, string MA001, string TA006)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERYS1 = new StringBuilder();
            StringBuilder SBQUERYS2 = new StringBuilder();

            if(!string.IsNullOrEmpty(MA001))
            {
                SBQUERYS1.AppendFormat(@" AND ( MA001 LIKE '%{0}%' OR MA002 LIKE '%{0}%' )", MA001);
            }
            else
            {
                SBQUERYS1.AppendFormat(@" ");
            }

            if (!string.IsNullOrEmpty(TA006))
            {
                SBQUERYS2.AppendFormat(@" AND ( TA006 LIKE '%{0}%' OR TA034 LIKE '%{0}%' )", TA006);
            }
            else
            {
                SBQUERYS2.AppendFormat(@" ");
            }

            SB.AppendFormat(@" 
                           
                                SELECT 
                                MA001 AS '廠商代號'
                                ,MA002 AS '廠商'
                                ,TA010 AS '預計到貨日'
                                ,TC001 AS '託外採購單別'
                                ,TC002 AS '託外採購單號'
                                ,TA006 AS '品號'
                                ,TA034 AS '品名'
                                ,TA015 AS '採購數量'
                                ,TA017 AS '進貨數量'
                                ,TA007 AS '單位'
                                ,TA001 AS '託外製令單別'
                                ,TA002 AS '託外製令單號'
                                ,TC045 
                                ,TA035 AS '規格'
                                FROM [TK].dbo.PURTC
                                LEFT JOIN [TK].dbo.MOCTA ON REPLACE(TA001+TA002,' ','')=TC045
                                LEFT JOIN [TK].dbo.PURMA ON MA001=TA032
                                WHERE TC001='A334' 
                                AND ISNULL(TC045,'')<>''
                                {0}
                                {1}

                            ", SBQUERYS1.ToString(), SBQUERYS2.ToString());


            return SB;

        }

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH(textBox1.Text, dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox2.Text.Trim(), textBox3.Text.Trim(), textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim());
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEARCH_MOCTO(textBox8.Text, dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(textBox9.Text.Trim(), textBox10.Text.Trim(), textBox15.Text.Trim(), textBox11.Text.Trim(), textBox12.Text.Trim(), textBox13.Text.Trim(), textBox14.Text.Trim());
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SEARCH_MOCTC(textBox16.Text, dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
        }
        private void button6_Click(object sender, EventArgs e)
        {
            string TA001 = textBox17.Text;
            string TA002 = textBox18.Text;
            string TC001 = "A334";
            string TC002;
            //string TC003 = textBox19.Text;
            string TC003 = dateTimePicker9.Value.ToString("yyyyMMdd");
            string TC024= dateTimePicker9.Value.ToString("yyyyMMdd");
            TC002 = GETMAXTC002(TC001, TC003);

            ADD_PURTC_PURTD(TA001,TA002,TC001, TC002, TC003, TC024);

            string TC045 = textBox17.Text.Trim() + textBox18.Text.Trim();
            //是否已產生託外採購單
            SERACH_PURTC(TC045);
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SEARCH_MOCTO_V2(textBox20.Text, dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string TO001 = textBox21.Text;
            string TO002 = textBox22.Text;
            string TO003 = textBox23.Text;
            string TE001 = textBox24.Text;
            string TE002 = textBox25.Text;
            string TE003;
            TE003 = GETMAXTE003(TE001, TE002);

            if (!string.IsNullOrEmpty(textBox24.Text) &&!string.IsNullOrEmpty(textBox25.Text))
            {

                ADD_PURTE_PURTF(TE001, TE002, TE003, TO001, TO002, TO003);

                string TE015 = TO001.Trim() + TO002.Trim() + TO003.Trim();
                //是否已產生託外採購單
                SERACH_PURTE(TE015);

            }
            else
            {
                MessageBox.Show("不是用外掛產生的採購單，無法再產生採購變更單");
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"),textBox26.Text.Trim(),textBox27.Text.Trim());
        }

        #endregion


    }
}
