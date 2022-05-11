using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKPUR
{
    public partial class FrmReportPURALL : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        public FrmReportPURALL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string TH004, string SDATE, string EDATE)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(TH004, SDATE, EDATE);
            Report report1 = new Report();
            report1.Load(@"REPORT\進貨採購請購表.frx");

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
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string TH004,string SDATE,string EDATE)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT TG003 AS '進貨日期',TH001 AS '進貨單別',TH002 AS '進貨單號',TH003 AS '進貨序號',TH004 AS '品號',TH005 AS '品名',TH007 AS '進貨數量',TH015 AS '驗收數量',TH016 AS '計價數量',TH017 AS '驗退數量',TH008 AS '進貨單位',TH011 AS '採購單別',TH012 AS '採購單號',TH013 AS '採購序號',ISNULL(TC003,'')  AS '採購日期',ISNULL(TD008,0)  AS '採購數量',ISNULL(TD009,'')  AS '採購單位',ISNULL(TA003,'')  AS '請購日期',ISNULL(TB001,'')  AS '請購單別',ISNULL(TB002,'')  AS '請購單號',ISNULL(TB003,'')  AS '請購序號',ISNULL(TB009,0)  AS '請購數量',ISNULL(TB007,'')  AS '請購單位',ISNULL(TA006,'') AS '請購單頭備註',ISNULL(TB012,'') AS '請購單身備註'
                            FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                            LEFT JOIN [TK].dbo.PURTD ON TD018='Y' AND TD001=TH011 AND TD002=TH012 AND TD003=TH013 AND TD004=TH004
                            LEFT JOIN [TK].dbo.PURTC ON TC014='Y' AND TC001=TD001 AND TC002=TD002 
                            LEFT JOIN [TK].dbo.PURTB ON TB025='Y' AND TB001=TD026 AND TB002=TD027 AND TB003=TD028 AND TB004=TH004
                            LEFT JOIN [TK].dbo.PURTA ON TA007='Y' AND TA001=TB001 AND TA002=TB002 

                            WHERE 1=1
                            AND TG013='Y'
                            AND TG001=TH001 AND TG002=TH002
                            AND (TH004 LIKE '%{0}%' OR TH005 LIKE '%{0}%')
                            AND TG003>='{1}' AND TG003<='{2}'
                            ORDER BY TH001,TH002,TH003

                            ", TH004, SDATE, EDATE);


            return SB;

        }

        public void SETFASTREPORT2(string TH012, string SDATE, string EDATE)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(TH012, SDATE, EDATE);
            Report report1 = new Report();
            report1.Load(@"REPORT\進貨採購請購表.frx");

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
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL2(string TH012, string SDATE, string EDATE)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT TG003 AS '進貨日期',TH001 AS '進貨單別',TH002 AS '進貨單號',TH003 AS '進貨序號',TH004 AS '品號',TH005 AS '品名',TH007 AS '進貨數量',TH015 AS '驗收數量',TH016 AS '計價數量',TH017 AS '驗退數量',TH008 AS '進貨單位',TH011 AS '採購單別',TH012 AS '採購單號',TH013 AS '採購序號',ISNULL(TC003,'')  AS '採購日期',ISNULL(TD008,0)  AS '採購數量',ISNULL(TD009,'')  AS '採購單位',ISNULL(TA003,'')  AS '請購日期',ISNULL(TB001,'')  AS '請購單別',ISNULL(TB002,'')  AS '請購單號',ISNULL(TB003,'')  AS '請購序號',ISNULL(TB009,0)  AS '請購數量',ISNULL(TB007,'')  AS '請購單位',ISNULL(TA006,'') AS '請購單頭備註',ISNULL(TB012,'') AS '請購單身備註'
                            FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                            LEFT JOIN [TK].dbo.PURTD ON TD018='Y' AND TD001=TH011 AND TD002=TH012 AND TD003=TH013 AND TD004=TH004
                            LEFT JOIN [TK].dbo.PURTC ON TC014='Y' AND TC001=TD001 AND TC002=TD002 
                            LEFT JOIN [TK].dbo.PURTB ON TB025='Y' AND TB001=TD026 AND TB002=TD027 AND TB003=TD028 AND TB004=TH004
                            LEFT JOIN [TK].dbo.PURTA ON TA007='Y' AND TA001=TB001 AND TA002=TB002 

                            WHERE 1=1
                            AND TG013='Y'
                            AND TG001=TH001 AND TG002=TH002
                            AND TH012 LIKE '%{0}%'
                            AND TG003>='{1}' AND TG003<='{2}'
                            ORDER BY TH001,TH002,TH003

                            ", TH012, SDATE, EDATE);


            return SB;

        }



        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(textBox1.Text.Trim(),dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(textBox2.Text.Trim(), dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
