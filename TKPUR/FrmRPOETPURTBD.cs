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
    public partial class FrmRPOETPURTBD : Form
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

        public FrmRPOETPURTBD()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\請購是否有採購.frx");

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
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (!string.IsNullOrEmpty(textBox1.Text))
            {
                STRQUERY.AppendFormat(@"  AND TB005 LIKE '%{0}%'", textBox1.Text);
            }
            else
            {
                STRQUERY.AppendFormat(@" ");
            }


            FASTSQL.AppendFormat(@"  SELECT PURMA.MA002 AS '廠商'");
            FASTSQL.AppendFormat(@"  ,TB001 AS '請購單別' ,TB002 AS '請購單號',TB003 AS '請購序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格'");
            FASTSQL.AppendFormat(@"  ,TB007 AS '請購單位',TB009 AS '請購數量',TB011 AS '需求日期',TB012 AS '備註'");
            FASTSQL.AppendFormat(@"  ,TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號'");
            FASTSQL.AppendFormat(@"  ,TD004 AS '採購品號',TD005 AS '採購品名',TD006 AS '採購規格',TD008 AS '採購數量'");
            FASTSQL.AppendFormat(@"  ,TD015 AS '已交數量',(TD008-TD015) AS '未交數量' ");
            FASTSQL.AppendFormat(@"  ,TD009 AS '單位',TD012 AS '預交日'");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.PURTA,[TK].dbo.PURTB ");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.VPURTD ON VTD013=TB001 AND VTD021=TB002 AND TD023=TB003    ");
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.PURMA ON TB010=MA001");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TB011>='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(STRQUERY.ToString());
            FASTSQL.AppendFormat(@"  ORDER BY TB011,TB010,TB004");           
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\請購未採明細表.frx");

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
            SQL = SETFASETSQL2();
            Table.SelectCommand = SQL;
            report1.Preview = previewControl2;
            report1.Show();

        }

        public string SETFASETSQL2()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

                     
            FASTSQL.AppendFormat(@"  
                                SELECT TA003 AS '請購日期',TB010 AS '廠商代',MA002 AS '廠商',TB001 AS '單別',TB002 AS '單號',TB003 AS '序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB009 AS '請購數量',TB007 AS '請購單位',TB039 AS '是否結案',TA012 AS '請購人代',MV002 AS '請購人',TB011 AS '需求日'
                                ,(SELECT ISNULL(SUM(TD008),0) FROM [TK].dbo.PURTC,[TK].dbo.PURTD WHERE TC001=TD001 AND TC002=TD002 AND TD026=TB001 AND TD027=TB002 AND TD028=TB003) AS '已採購數量'
                                FROM [TK].dbo.PURTA,[TK].dbo.CMSMV,[TK].dbo.PURTB
                                LEFT JOIN [TK].dbo.PURMA ON  TB010=MA001

                                WHERE TA001=TB001 AND TA002=TB002
                                AND TA012=MV001
                                AND TB009>0
                                AND TB025='Y'
                                AND TB039='N' 
                                AND (TB009-(SELECT ISNULL(SUM(TD008),0) FROM [TK].dbo.PURTC,[TK].dbo.PURTD WHERE TC001=TD001 AND TC002=TD002 AND TD026=TB001 AND TD027=TB002 AND TD028=TB003))>0
                                ORDER BY TA003,TB001,TB002,TB003
                                
                                ");

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();

            textBox1.Text = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();

        }
        #endregion

    }
}
