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
            FASTSQL.AppendFormat(@"  LEFT JOIN [TK].dbo.PURTD ON TD004=TB004 AND TD023=TB003");
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
