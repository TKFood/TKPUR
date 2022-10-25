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

namespace TKPUR
{
    public partial class frmREPORTINVLAUSED : Form
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

        public frmREPORTINVLAUSED()
        {
            InitializeComponent();

            comboBox1load();


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

                                SELECT [ID],[KINDS]
                                FROM [TKPUR].[dbo].[KINDS]
                                WHERE [TYPES]='PUR'
                                ORDER BY ID
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("KINDS", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "KINDS";
            comboBox1.DisplayMember = "KINDS";
            sqlConn.Close();


        }
        public void SETFASTREPORT(string SDATES,string EDATES,string MB001,string KINDS)
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\品號進銷領表.frx");

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

            SQL = SETFASETSQL(SDATES, EDATES, MB001, KINDS);

            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string SDATES, string EDATES, string MB001,string KINDS)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if(KINDS.Equals("全部"))
            {
                STRQUERY.AppendFormat(@" ");
            }
            else
            {
                STRQUERY.AppendFormat(@" AND NEWMQ008='{0}'", KINDS);
            }          
            FASTSQL.AppendFormat(@" 
                                    SELECT NEWMQ008 AS '分類',LA001 AS '品號',SUBSTRING(LA004,1,6)  AS '年月',SUM(LA005*LA011)  AS '數量',INVMB.MB002 AS '品名',INVMB.MB003 AS '規格',INVMB.MB004 AS '單位'
                                    ,INVMB.MB050 AS '最近進貨價'
                                    
                                    FROM 
                                    (
                                    SELECT MQ001,MQ002,MQ008,LA001,LA004,LA005,LA006,LA007,LA008,LA011,MB002,MB003,MB004
                                    ,CASE  WHEN MQ001 IN ('A421','A422','A431') AND LA005=-1 THEN '4組合領用' WHEN MQ001 IN ('A421','A422','A431') AND LA005=1 THEN '5組合生產'  WHEN MQ008='1' THEN '1進貨/入庫'  WHEN MQ008='2' THEN '2銷貨'  WHEN MQ008='3' THEN '3領用' END NEWMQ008 
                                    FROM [TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.CMSMQ WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK)
                                    WHERE LA006=MQ001
                                    AND LA001=MB001
                                    AND LA004>='{0}' AND LA004<='{1}'
                                    AND MQ008 IN ('','1','2','3')
                                    AND (LA001 LIKE  '%{2}%' OR MB002 LIKE '%{2}%')
                                    ) AS TEMP,[TK].dbo.INVMB
                                    WHERE LA001=MB001
                                    {3}
                                    GROUP BY LA001,NEWMQ008,SUBSTRING(LA004,1,6),INVMB.MB002,INVMB.MB003,INVMB.MB004,INVMB.MB050
                                    ORDER BY LA001,NEWMQ008,SUBSTRING(LA004,1,6),INVMB.MB002,INVMB.MB003,INVMB.MB004,INVMB.MB050
                                    ",SDATES,EDATES,MB001, STRQUERY.ToString());

            return FASTSQL.ToString();
        }

        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox1.Text,comboBox1.Text.ToString());
            }
            
        }

        #endregion
    }
}
