﻿using System;
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
            // 取得今年的最後一天
            DateTime lastDayOfMonth = DateTime.Now;

            dateTimePicker1.Value = firstDayOfMonth;
            dateTimePicker2.Value = lastDayOfMonth;
            dateTimePicker3.Value = firstDayOfMonth;
            dateTimePicker4.Value = lastDayOfMonth;
            dateTimePicker5.Value = firstDayOfMonth;
            dateTimePicker6.Value = lastDayOfMonth;
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
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox17.Text = row.Cells["單別"].Value.ToString().Trim();
                    textBox18.Text = row.Cells["單號"].Value.ToString().Trim();
                    textBox19.Text = row.Cells["TA003"].Value.ToString().Trim();
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
            string TC001 = "A334";
            string TC002;
            string TC003 = textBox19.Text;
            TC002 = GETMAXTC002(TC001, TC003);

        }

        #endregion


    }
}
