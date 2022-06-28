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
using TKITDLL;


namespace TKPUR
{
    public partial class FrmREPORTPURTL : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds4 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        Report report1 = new Report();

        public FrmREPORTPURTL()
        {
            InitializeComponent();

            dateTimePicker1.Value = Convert.ToDateTime(DateTime.Today.Year + "/1/1");
        }

        #region FUNCTION
        public void SETFASTREPORT(string TL003, string TL004)
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\核價單廠商調價表.frx");

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
            SQL = SETFASETSQL(TL003, TL004);

            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string TL003,string TL004)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();


            FASTSQL.AppendFormat(@"  
                                 SELECT TL004 AS '廠商代號',MA002 AS '廠商',TL003 AS '核價日期',TL005 AS '幣別',TM004 AS '品號',TM005 AS '品名',TM006 AS '規格',TM010 AS '單價',TM009 AS '計價單位',TM008 AS '分量計價',TM014 AS '生效日',TM015 AS '失效日',TN007 AS '數量以上',TN008 AS '分量計價單價'
                                    FROM [TK].dbo.PURMA,[TK].dbo.PURTL, [TK].dbo.PURTM  
                                    LEFT JOIN [TK].dbo.PURTN ON TM001=TN001 AND TM002=TN002 AND TM003=TN003  
                                    WHERE 1=1
                                    AND MA001=TL004
                                    AND TL001=TM001 AND TL002=TM002
                                    AND TL006='Y'
                                    AND TL003>='{0}'
                                    AND (TL004 LIKE '{1}%' OR MA002 LIKE '{1}%')
                                    ORDER BY TL004,MA002,TM004,TL003,TM009
 

                                ", TL003, TL004);

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"),textBox1.Text.Trim());
        }
        #endregion

    }
}
