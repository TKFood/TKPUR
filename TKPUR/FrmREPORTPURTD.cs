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

namespace TKPUR
{
    public partial class FrmREPORTPURTD : Form
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

        public FrmREPORTPURTD()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\預計採購表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
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

            if(!string.IsNullOrEmpty(textBox1.Text))
            {
                STRQUERY.AppendFormat(@"  AND TD005 LIKE '%{0}%'", textBox1.Text);
            }
            else
            {
                STRQUERY.AppendFormat(@"  ");
            }
          
     
            FASTSQL.AppendFormat(@"  SELECT MA002 AS '廠商',TD004 AS '品號', TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量',TD015 AS '已交量',TD009 AS '單位',TD010  AS '單價',TD011  AS '金額',TD012 AS '預交日',TD014 ");
            FASTSQL.AppendFormat(@"  ,(SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') AS COMMENT1");
            FASTSQL.AppendFormat(@"  ,CASE WHEN ISNULL(TD014,'')<>'' THEN TD014 ELSE (SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') END AS '備註'");
            FASTSQL.AppendFormat(@"  ,TD001,TD002,TD003");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA ");
            FASTSQL.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
            FASTSQL.AppendFormat(@"  AND MA001=TC004");
            FASTSQL.AppendFormat(@"  AND TD012>='{0}' AND TD012<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            FASTSQL.AppendFormat(@"  AND TD018='Y'");
            FASTSQL.AppendFormat(STRQUERY.ToString());
            FASTSQL.AppendFormat(@"  ORDER BY CASE WHEN ISNULL(TD014,'')<>'' THEN TD014 ELSE (SELECT TOP 1 TD014 FROM [TK].dbo.PURTD A WHERE A.TD001=PURTD.TD001 AND A.TD002=PURTD.TD002 AND A.TD003='0001') END");
            FASTSQL.AppendFormat(@"  ,TD012,TD001,TD002,TD003");
            FASTSQL.AppendFormat(@"  ");

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();

            textBox1.Text = null;
        }
        #endregion

    }
}
