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
    public partial class FrmREPURPLAN : Form
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

        public FrmREPURPLAN()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\原料採購計畫.frx");

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

 
            SB.AppendFormat(@" 
                            SELECT 品號,品名,單位,單價,期初存貨,期末存貨,本期秏用數量,(期末存貨+本期秏用數量-期初存貨) AS '本期採購數',(期末存貨+本期秏用數量-期初存貨)*單價 AS '金額'
                            FROM (
                            SELECT MB001 AS '品號',MB002 AS '品名',MB004 AS '單位',MB050 AS '單價'
                            ,ISNULL((SELECT SUM(LA011*LA005)  FROM [TK].dbo.INVLA WITH(NOLOCK)  WHERE LA001=MB001 and LA004<'{0}' ),0) AS '期初存貨'
                            ,ISNULL((SELECT SUM(LA011*LA005)  FROM [TK].dbo.INVLA  WITH(NOLOCK) WHERE LA001=MB001 and LA004<'{1}' ),0) AS '期末存貨'
                            ,ISNULL((SELECT SUM(LA011*LA005)*-1  FROM [TK].dbo.INVLA WITH(NOLOCK)  WHERE LA001=MB001 AND LA005=-1 AND LA004>='{0}' AND LA004<'{1}'),0) AS '本期秏用數量'
                            FROM [TK].dbo.INVMB  WITH(NOLOCK)
                            WHERE  MB001 LIKE '1%'
                            AND MB002 NOT LIKE '%暫停%'
                            ) AS TEMP
                            ORDER BY 品號

                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
         

            return SB;

        }

        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\物料採購計畫.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report2.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

                        
            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@" 
                            SELECT 類別,單位,AVG(單價) AS '單價',SUM(期初存貨) AS '期初存貨',SUM(期末存貨) AS '期末存貨',SUM(本期秏用數量) AS '本期秏用數量',SUM((期末存貨+本期秏用數量-期初存貨)) AS '本期採購數',SUM((期末存貨+本期秏用數量-期初存貨)*單價) AS '金額'
                            FROM (
                            SELECT MA003 AS '類別', MB001 AS '品號',MB002 AS '品名',MB004 AS '單位',MB050 AS '單價'
                            ,ISNULL((SELECT SUM(LA011*LA005)  FROM [TK].dbo.INVLA  WITH(NOLOCK) WHERE LA001=MB001 and LA004<'{0}' ),0) AS '期末存貨'
                            ,ISNULL((SELECT SUM(LA011*LA005)  FROM [TK].dbo.INVLA WITH(NOLOCK)  WHERE LA001=MB001 and LA004<'{1}' ),0) AS '期初存貨'
                            ,ISNULL((SELECT SUM(LA011*LA005)*-1  FROM [TK].dbo.INVLA WITH(NOLOCK)  WHERE LA001=MB001 AND LA005=-1 AND LA004>='{0}' AND LA004<'{1}'),0) AS '本期秏用數量'
                            FROM [TK].dbo.INVMB,[TK].dbo.INVMA
                            WHERE  MA001='5'
                            AND MB111=MA002
                            AND MB001 LIKE '2%'
                            AND MB002 NOT LIKE '%暫停%' ) AS TEMP
                            GROUP BY 類別,單位
                            ORDER BY 類別,單位
                            ", dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));

            return SB;

        }

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }

        #endregion


    }
}
