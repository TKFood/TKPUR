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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;
using NPOI.SS.Formula.Functions;
namespace TKPUR
{
    public partial class frmREPORTINCOUNTS : Form
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
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;

        public Report report1 { get; private set; }

        public frmREPORTINCOUNTS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
        }
        #region FUNCTION
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

                               SELECT  [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='frmREPORTINCOUNTS'
                                ORDER BY [ID]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("PARAID", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
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

                               SELECT  [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='frmREPORTINCOUNTSKIND'
                                ORDER BY [ID]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("PARAID", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARAID";
            comboBox2.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void SETFASTREPORT(string SDATES, string EDATES, string KINDS,string MB001,string QUERYKINDS)
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\進貨排名表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;

            SQL = SETFASETSQL(SDATES, EDATES, KINDS, MB001, QUERYKINDS);

            Table.SelectCommand = SQL;

            report1.SetParameterValue("P1", SDATES);
            report1.SetParameterValue("P2", EDATES);

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string SDATES, string EDATES, string KINDS,string MB001,string QUERYKINDS)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            StringBuilder STRQUERY2 = new StringBuilder();

            if (!string.IsNullOrEmpty(MB001))
            { 
                STRQUERY.AppendFormat(@" AND (TH004 LIKE '%{0}%' OR TH005 LIKE '%{0}%')", MB001);
            }
            else
            {
                STRQUERY.AppendFormat(@"");
            }
            if(QUERYKINDS.Equals("原料"))
            {
                STRQUERY2.AppendFormat(@" AND TH004 LIKE '1%' ");
            }
            else if (QUERYKINDS.Equals("物料"))
            {
                STRQUERY2.AppendFormat(@" AND TH004 LIKE '2%' ");
            }
            else
            {
                STRQUERY2.AppendFormat(@"");
            }

            if (KINDS.Equals("依數量"))
            {
                FASTSQL.AppendFormat(@" 
                                    SELECT ROW_NUMBER ()  OVER(ORDER BY SUM(LA011) DESC) AS 'SERNO'
                                    ,TH004 AS '品號',TH005 AS '品名',SUM(TH047) AS '進貨未稅金額',SUM(LA011) AS '進貨數量',MB004 AS '單位'
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.INVLA,[TK].dbo.INVMB 
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                                    AND TH004=MB001
                                    AND TG013 IN ('Y')
                                    AND TG003>='{0}' AND TG003<='{1}'
                                    {2}
                                    {3}

                                    GROUP BY TH004,TH005,MB004
                                    ORDER BY SUM(LA011) DESC
                                    ", SDATES, EDATES, STRQUERY.ToString(), STRQUERY2.ToString());
            }
            else if (KINDS.Equals("依金額"))
            {
                FASTSQL.AppendFormat(@" 
                                    SELECT ROW_NUMBER ()  OVER(ORDER BY SUM(TH047) DESC) AS 'SERNO'
                                    ,TH004 AS '品號',TH005 AS '品名',SUM(TH047) AS '進貨未稅金額',SUM(LA011) AS '進貨數量',MB004 AS '單位'
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.INVLA,[TK].dbo.INVMB 
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                                    AND TH004=MB001
                                    AND TG013 IN ('Y')
                                    AND TG003>='{0}' AND TG003<='{1}'
                                    {2}
                                    {3}
                                    GROUP BY TH004,TH005,MB004
                                    ORDER BY SUM(TH047) DESC
                                    ", SDATES, EDATES, STRQUERY.ToString(), STRQUERY2.ToString());
            }
          

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            string KINDS = comboBox1.Text.ToString();
            string MB001=textBox1.Text.ToString().Trim();
            string QUERYKINDS= comboBox2.Text.ToString();
            SETFASTREPORT(SDATES, EDATES, KINDS, MB001, QUERYKINDS);
        }

        #endregion
    }
}
