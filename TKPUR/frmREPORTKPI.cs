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
    public partial class frmREPORTKPI : Form
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


        public frmREPORTKPI()
        {
            InitializeComponent();

            
        }

        #region FUNCTION
        private void frmREPORTKPI_Load(object sender, EventArgs e)
        {
            comboBox1load();

            SETDATES();
        }
        public void SETDATES()
        {
            // 取得今年的第一天
            DateTime firstDayOfYear = new DateTime(DateTime.Now.Year, 1, 1);
            // 取得今年的最後一天
            DateTime lastDayOfYear = new DateTime(DateTime.Now.Year, 12, 31);

            dateTimePicker2.Value = firstDayOfYear;
            dateTimePicker3.Value = lastDayOfYear;
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
                                SELECT
                                [ID]
                                ,[KINDS]
                                ,[REPORTNAMES]
                                FROM [TKPUR].[dbo].[TBREPOSRTNAMES]
                                WHERE [KINDS]='PUR'
                                ORDER BY [ID]
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("REPORTNAMES", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "REPORTNAMES";
            comboBox1.DisplayMember = "REPORTNAMES";
            sqlConn.Close();


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.Text.Equals("應付帳款及進貨佔比"))
            {
                dateTimePicker1.Enabled = true;

                dateTimePicker2.Enabled = false;
                dateTimePicker3.Enabled = false;
            }
            else if(comboBox1.Text.Equals("新進廠商清單"))
            {
                dateTimePicker1.Enabled = false;

                dateTimePicker2.Enabled = true;
                dateTimePicker3.Enabled = true;
            }
        }
        public void SETFASTREPORT(string REPORTNAME)
        {
            report1 = new Report();
            string SQL="";

            

            if (REPORTNAME.Equals("應付帳款及進貨佔比"))
            {
                report1.Load(@"REPORT\採購指標.frx");
                DateTime YEARFIRSTDAYS = new DateTime(dateTimePicker1.Value.Year, 1, 1);
                SQL = SETFASETSQL1(YEARFIRSTDAYS.ToString("yyyyMMdd"));
            }
            else if(REPORTNAME.Equals("新進廠商清單"))
            {
                report1.Load(@"REPORT\廠商清單.frx");

                SQL = SETFASETSQL2(dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P1", dateTimePicker2.Value.ToString("yyyyMMdd"));
                report1.SetParameterValue("P2", dateTimePicker3.Value.ToString("yyyyMMdd"));
            }


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

            
            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL1(string YEARFIRSTDAYS)
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"   
                                WITH MonthData AS (
                                    SELECT CAST(YEAR('{0}') AS VARCHAR(4)) AS YEARS,
                                           1 AS Month,
		                                   RIGHT('0' + CAST(1 AS VARCHAR(2)), 2) AS MONTHS,
		                                   RIGHT('0' + CAST(2 AS VARCHAR(2)), 2) AS NEXTMONTHS,
		                                   CAST(YEAR('{0}') AS VARCHAR(4)) AS NEXTYEARS
                                    UNION ALL
                                    SELECT YEARS,
                                           Month + 1,
		                                   RIGHT('0' + CAST((Month + 1) AS VARCHAR(2)), 2) AS MONTHS,
		                                   CASE WHEN (Month + 2)<>13 THEN RIGHT('0' + CAST((Month + 2) AS VARCHAR(2)), 2) ELSE '01' END AS NEXTMONTHS,
		                                   CASE WHEN (Month + 2)<>13 THEN  CAST(YEAR('{0}') AS VARCHAR(4))  ELSE  CAST((YEAR('{0}')+1) AS VARCHAR(4)) END AS NEXTYEARS
                                    FROM MonthData
                                    WHERE Month < 12
                                )

                                SELECT *
                                ,(應付款款期初一月+應付款款期末) AS '應付帳款區間小計'
                                ,((應付款款期初一月+應付款款期末)/2) AS '平均應付帳款'
                                ,(銷貨成本累計/((應付款款期初一月+應付款款期末)/2)) AS '應付帳款周轉率'
                                ,CASE WHEN (銷貨成本累計/((應付款款期初一月+應付款款期末)/2)) >0 THEN (累積天數/(銷貨成本累計/((應付款款期初一月+應付款款期末)/2)) ) ELSE 0 END  AS '應付帳款周轉天數'
                                ,(CASE WHEN 進貨各月總金額>0 AND 銷貨各月金額>0 THEN 進貨各月總金額/銷貨各月金額 ELSE 0 END ) AS '進貨佔營收佔比'
                                FROM 
                                (
                                SELECT YEARS,MONTHS,NEXTMONTHS,NEXTYEARS
                                ,(
                                SELECT ISNULL(SUM(LA017-LA020-LA022-LA023),0)
                                FROM [TK].dbo.SASLA
                                WHERE YEAR(LA015)=YEARS AND MONTH(LA015)=MONTHS
                                ) AS '銷貨各月金額'
                                ,(
                                SELECT ISNULL(SUM(LA024),0)
                                FROM [TK].dbo.SASLA
                                WHERE YEAR(LA015)=YEARS AND MONTH(LA015)=MONTHS
                                ) AS '銷貨各月成本'
                                ,(
                                SELECT ISNULL(SUM(LA024),0)
                                FROM [TK].dbo.SASLA
                                WHERE YEAR(LA015)=YEARS AND MONTH(LA015)<=MONTHS
                                ) AS '銷貨成本累計'
                                ,(
                                SELECT ISNULL(SUM(TH047+TH048),0) AS TOTALMONEYS
                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                WHERE 1=1
                                AND TG001=TH001 AND TG002=TH002
                                AND TG013='Y'
                                AND TG003>=YEARS+MONTHS+'01'
                                AND TG003<=YEARS+MONTHS+'31'
                                ) AS '進貨各月總金額'
                                ,
                                (
                                SELECT ISNULL(SUM(TA028+TA029),0)
                                FROM [TK].dbo.ACPTA
                                WHERE 1=1
                                AND TA024='Y'
                                AND TA003<YEARS+MONTHS+'01'
                                AND TA051>YEARS+MONTHS+'01'
                                ) AS '應付款款期初'
                                ,
                                (
                                SELECT ISNULL(SUM(TA028+TA029),0)
                                FROM [TK].dbo.ACPTA
                                WHERE 1=1
                                AND TA024='Y'
                                AND TA003<NEXTYEARS+NEXTMONTHS+'01'
                                AND TA051>NEXTYEARS+NEXTMONTHS+'01'
                                ) AS '應付款款期末'
                                ,
                                (
                                SELECT SUM(TA028+TA029)
                                FROM [TK].dbo.ACPTA
                                WHERE 1=1
                                AND TA024='Y'
                                AND TA003<YEARS+'0101'
                                AND TA051>YEARS+'0101'
                                ) AS '應付款款期初一月'
                                ,CASE WHEN NEXTYEARS=YEARS THEN DATEDIFF(day, YEARS+'0101', YEARS+NEXTMONTHS+'01') ELSE DATEDIFF(day, YEARS+'0101', NEXTYEARS+NEXTMONTHS+'01') END AS '累積天數'
                                FROM MonthData
                                ) AS TEMP                               

                                ", YEARFIRSTDAYS);

            return FASTSQL.ToString();
        }

        public string SETFASETSQL2(string SDAY,string EDAY)
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"                                   
                                SELECT 
                                MA001 AS '廠商代號'
                                ,MA002 AS '廠商'
                                ,MA014 AS '地址'
                                ,MA005 AS '統一編號'
                                ,CREATE_DATE AS '建立日期'
                                FROM [TK].dbo.PURMA
                                WHERE CREATE_DATE>='{0}'
                                AND CREATE_DATE<='{1}'
                                ORDER BY MA001

                                ", SDAY, EDAY);

            return FASTSQL.ToString();
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
