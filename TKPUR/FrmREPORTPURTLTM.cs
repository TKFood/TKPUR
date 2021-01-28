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

namespace TKPUR
{
    public partial class FrmREPORTPURTLTM : Form
    {

        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        DataSet ds = new DataSet();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();

        public Report report1 { get; private set; }

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public FrmREPORTPURTLTM()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
           
        }
        public void SETFASTREPORT()
        {
            string SQL;
            string SQL2;
            report1 = new Report();
            report1.Load(@"REPORT\原料、物料核價單漲跌整理表.frx");

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

            FASTSQL.AppendFormat(@"   
                                SELECT TL003 AS '核價日',MA002 AS '廠商',TL004 AS '廠商ID',TM004 AS '品號',TM005 AS '品名',TM010 AS '單價',TL001 AS '核價單別',TL002 AS '核價單號'
                                ,(SELECT TOP 1 TL003 FROM [TK].dbo.PURTL TL,[TK].dbo.PURTM TM WHERE TL.TL001=TM.TM001 AND TL.TL002=TM.TM002 AND TM.TM004=PURTM.TM004 AND TL.TL004=PURTL.TL004 AND TL.TL003<>PURTL.TL003 ORDER BY TL003 DESC) AS '上次核價日'
                                ,(SELECT TOP 1 TM010 FROM [TK].dbo.PURTL TL,[TK].dbo.PURTM TM WHERE TL.TL001=TM.TM001 AND TL.TL002=TM.TM002 AND TM.TM004=PURTM.TM004 AND TL.TL004=PURTL.TL004 AND TL.TL003<>PURTL.TL003 ORDER BY TL003 DESC) AS '上次核價單價'

                                ,(SELECT TOP 1 TL003+'-'+CONVERT(NVARCHAR,TM010) FROM [TK].dbo.PURTL TL,[TK].dbo.PURTM TM WHERE TL.TL001=TM.TM001 AND TL.TL002=TM.TM002 AND TM.TM004=PURTM.TM004 AND TL.TL004=PURTL.TL004 AND TL.TL003<>PURTL.TL003 ORDER BY TL003 DESC) AS '備註'
                                FROM [TK].dbo.PURTL,[TK].dbo.PURTM,[TK].dbo.PURMA
                                WHERE TL001=TM001 AND TL002=TM002
                                AND MA001=TL004
                                AND TL006='Y'
                                AND TL003>='{0}' AND TL003<='{1}'
                                ORDER BY MA002,TL003,TM004

                                ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

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
