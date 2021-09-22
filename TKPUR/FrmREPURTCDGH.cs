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
    public partial class FrmREPURTCDGH : Form
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

        public FrmREPURTCDGH()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\原物料交期準時表.frx");

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

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT 廠商,品號,品名,規格,需求量,單位,需求日,採購單別,採購單號,採購序號,預交日到貨量,已到貨量,最後到貨日");
            SB.AppendFormat(" ,(CASE WHEN 需求量>預交日到貨量 THEN '少交' ELSE '' END ) AS '預交日數量狀態'");
            SB.AppendFormat(" ,(CASE WHEN 需求量>已到貨量 THEN '少交' ELSE '' END ) AS '到貨數量狀態'");
            SB.AppendFormat(" ,(CASE WHEN (需求日<最後到貨日 OR ISNULL(最後到貨日,'')='') THEN '遲交' ELSE '' END ) AS '日期狀態'");
            SB.AppendFormat(" FROM (");
            SB.AppendFormat(" SELECT MA002 AS '廠商',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '需求量',TD009 AS '單位',TD012 AS '需求日',TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH007),0) FROM [TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TH011=TD001 AND TH012=TD002 AND TH013=TD003 AND TG003<=TD012) AS '預交日到貨量'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH007),0) FROM [TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TH011=TD001 AND TH012=TD002 AND TH013=TD003 ) AS '已到貨量'");
            SB.AppendFormat(" ,(SELECT TOP 1 TG003  FROM [TK].dbo.PURTH,[TK].dbo.PURTG WHERE TG001=TH001 AND TG002=TH002 AND TH011=TD001 AND TH012=TD002 AND TH013=TD003 ORDER BY TG001,TG002 DESC) AS '最後到貨日'");
            SB.AppendFormat(" FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TC004=MA001");
            SB.AppendFormat(" AND TD012>='{0}' AND TD012<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" ORDER BY 廠商,品號,需求日");
            SB.AppendFormat("  ");
            SB.AppendFormat(" ");

            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
