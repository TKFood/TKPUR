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
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\原物料張跌表.frx");

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
                            SELECT 
                            TH004,TH005,TH018,TH008
                            ,(SELECT SUM(TH007) FROM [TK].dbo.PURTG TG ,[TK].dbo.PURTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002 AND  TG.TG003>='{0}' AND TG.TG003<='{1}' AND TH.TH004=TEMP.TH004 AND TH.TH008=TEMP.TH008 ) AS 'NUM'
                            ,(SELECT TOP 1 SUBSTRING(TG003, 1, 6) FROM[TK].dbo.PURTG TG ,[TK].dbo.PURTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002= TH.TH002 AND TG003>='{0}' AND TG003<='{1}' AND TH.TH004= TEMP.TH004 AND TH.TH018= TEMP.TH018) AS 'YM'
                            FROM(
                            SELECT TH004, TH005, TH018, TH008
                            FROM [TK].dbo.PURTG, [TK].dbo.PURTH
                            WHERE TG001= TH001 AND TG002 = TH002
                            AND ( TH004 LIKE '1%' OR TH004 LIKE '2%')
                            AND TH004 NOT LIKE '199%'
                            AND TH004 NOT LIKE '299%'
                            AND TG003>='{0}' AND TG003<='{1}'
                            GROUP BY  TH004,TH005,TH018,TH008
                            ) AS TEMP
                            WHERE TH004 IN(
                            SELECT TH004
                            FROM (
                            SELECT TH004, TH005, TH018, TH008
                            FROM [TK].dbo.PURTG, [TK].dbo.PURTH
                            WHERE TG001= TH001 AND TG002 = TH002
                            AND ( TH004 LIKE '1%' OR TH004 LIKE '2%')
                            AND TH004 NOT LIKE '199%'
                            AND TH004 NOT LIKE '299%'
                            AND TG003>='{0}' AND TG003<='{1}'
                            GROUP BY  TH004,TH005,TH018,TH008
                            ) AS TEMP
                            GROUP BY TH004,TH005,TH008
                            HAVING COUNT(TH004)>=2
                            )ORDER BY TH004,(SELECT TOP 1 SUBSTRING(TG003,1,6) FROM[TK].dbo.PURTG TG,[TK].dbo.PURTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002= TH.TH002 AND TG003>='{0}' AND TG003<='{1}' AND TH.TH004= TEMP.TH004 AND TH.TH018= TEMP.TH018)
                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

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
