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
        private void FrmREPURPRICE_Load(object sender, EventArgs e)
        {
            SETDATE();
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
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='FrmREPURPRICE'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARAID";
            comboBox1.DisplayMember = "PARAID";
            sqlConn.Close();


        }
        public void SETDATE()
        {
            dateTimePicker1.Value = new DateTime(DateTime.Now.Year, 1, 1);
            dateTimePicker2.Value = new DateTime(DateTime.Now.Year, 12, 31);
        }
        public void SETFASTREPORT(string SDATES,string EDATES,string TH004,string KIND)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDATES, EDATES, TH004, KIND);
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

        public StringBuilder SETSQL(string SDATES, string EDATES, string TH004,string KIND)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();
            StringBuilder SBQUERY99 = new StringBuilder();

            if (!string.IsNullOrEmpty(TH004))
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

            if (!string.IsNullOrEmpty(KIND))
            {
                if (KIND.Equals("原料"))
                {
                    SBQUERY99.AppendFormat(@"  AND (TH004 LIKE '1%' )  ");
                }
                else if (KIND.Equals("物料"))
                {
                    SBQUERY99.AppendFormat(@"  AND (TH004 LIKE '2%' )  ");
                }
                else if (KIND.Equals("其他"))
                {
                    SBQUERY99.AppendFormat(@"  AND TH004 NOT LIKE '1%' 
                                                AND TH004 NOT LIKE '2%'  ");
                }
                else if (KIND.Equals("全部"))
                {
                    SBQUERY99.AppendFormat(@"   ");
                }
            }
            else
            {
                SBQUERY99.AppendFormat(@"
                                       
                                        ");
            }

            SB.AppendFormat(@"       
                            WITH MonthlyData AS (
                                SELECT 
                                    SUBSTRING(TG003,1,6) AS YM,
                                    TH004 AS 品號,
                                    TH005 AS 品名,
                                    TH008 AS 單位,
                                    SUM(TH007) AS 進貨數量,
                                    SUM(TH016) AS 計價數量,
                                    SUM(TH047+TH048) AS 本幣金額,
                                    (CASE 
                                        WHEN SUM(TH047+TH048) > 0 AND SUM(TH016) > 0 
                                        THEN CONVERT(decimal(16,3),SUM(TH047+TH048) / SUM(TH016) )
                                        ELSE 0 
                                    END) AS 進貨單價
		                            ,(SELECT SUM(TH007) FROM [TK].dbo.PURTG TG ,[TK].dbo.PURTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002 AND  TG.TG003>='20240101' AND TG.TG003<='20250228' AND TH.TH004=PURTH.TH004 AND TH.TH008=PURTH.TH008 ) AS '進貨總數量'
                                FROM [TK].dbo.PURTG, [TK].dbo.PURTH
                                WHERE TG001 = TH001
                                AND TG002 = TH002
                                AND TG013 = 'Y'
                                AND TH004 NOT LIKE '199%'
                                AND TH004 NOT LIKE '299%'
                                AND TG003 >= '{0}' AND TG003 <= '{1}'
                                {2}
                                {3}
                                GROUP BY SUBSTRING(TG003,1,6), TH004, TH005, TH008
                            )
                            SELECT A.YM, A.品號, A.品名, A.單位,A.進貨總數量, A.進貨單價,A.進貨數量,A.計價數量,
                                   B.進貨單價 AS 前月單價,
                                   CASE 
                                       WHEN A.進貨單價 > B.進貨單價 THEN '↑ 上漲'
                                       WHEN A.進貨單價 < B.進貨單價 THEN '↓ 下跌'
                                       ELSE '→ 相同'
                                   END AS 單價變化
                            FROM MonthlyData A
                            LEFT JOIN MonthlyData B 
                                ON A.品號 = B.品號 
                                AND A.YM = SUBSTRING(CONVERT(VARCHAR(8), DATEADD(MONTH, 1, CAST(B.YM + '01' AS DATE)), 112), 1, 6)
                            ORDER BY A.品號, A.YM;

                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), SBQUERY1.ToString(), SBQUERY99.ToString());

            return SB;
             
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), textBox1.Text.Trim(), comboBox1.Text);
        }
        #endregion

       
    }
}
