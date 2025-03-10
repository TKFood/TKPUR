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
    public partial class FrmREPURPRICE : Form
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

        public FrmREPURPRICE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string SDATES,string EDATES,string TH004)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES, TH004);
            Report report1 = new Report();
            report1.Load(@"REPORT\原物料漲跌表V2.frx");

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

        public StringBuilder SETSQL(string SDATES, string EDATES, string TH004)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();

            if(!string.IsNullOrEmpty(TH004))
            {
                SBQUERY1.AppendFormat(@"
                                        AND (TH004 LIKE '%{0}%' OR TH005 LIKE '%{0}%')
                                        ", TH004);
            }
            else
            {
                SBQUERY1.AppendFormat(@"
                                       
                                        ");
            }
            
            SB.AppendFormat(@"                            
                                SELECT 
                                SUBSTRING(TG003,1,6) AS 'YM'
                                ,TH004 AS '品號'
                                ,TH005 AS '品名'
                                ,TH008 AS '單位'
                                ,SUM(TH007) AS '進貨數量'
                                ,SUM(TH016) AS '計價數量'
                                ,SUM(TH047+TH048) AS '本幣金額'
                                ,(CASE WHEN SUM(TH047+TH048)>0 AND SUM(TH016)>0 THEN SUM(TH047+TH048)/SUM(TH016) ELSE 0 END )  AS '進貨單價'
                                ,(SELECT SUM(TH007) FROM [TK].dbo.PURTG TG ,[TK].dbo.PURTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002 AND  TG.TG003>='20240101' AND TG.TG003<='20250228' AND TH.TH004=PURTH.TH004 AND TH.TH008=PURTH.TH008 ) AS '進貨總數量'

                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                WHERE TG001=TH001
                                AND TG002=TH002
                                AND TG013='Y'
                                AND TH004 NOT LIKE '199%'
                                AND TH004 NOT LIKE '299%'
                                AND TG003>='{0}' AND TG003<='{1}'
                                {2}

                                GROUP BY SUBSTRING(TG003,1,6) ,TH004,TH005,TH008
                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), SBQUERY1.ToString());

            return SB;

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"),textBox1.Text.Trim());
        }
        #endregion
    }
}
