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
    public partial class FrmREPORTINLVA : Form
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


        public FrmREPORTINLVA()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\進貨成本佔成本比率分析表.frx");

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

            FASTSQL.AppendFormat(@"  SELECT SEQ,KIND,MONTHS,MONEYS");
            FASTSQL.AppendFormat(@"  FROM (");
            FASTSQL.AppendFormat(@"  SELECT '1A' AS SEQ ,'營收' AS KIND,SUBSTRING(TA003,1,6) AS MONTHS,SUM(TB004*TB007)*-1 AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.ACTTA  WITH (NOLOCK),[TK].dbo.ACTTB WITH (NOLOCK)");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TA003 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND TB005 LIKE '4%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(TA003,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '1B' AS SEQ ,'成本率' AS KIND,SUBSTRING(TA003,1,6) AS MONTHS ,0 AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.ACTTA  WITH (NOLOCK),[TK].dbo.ACTTB WITH (NOLOCK)");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TA003 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND TB005 LIKE '4%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(TA003,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '1C' AS SEQ ,'成本額' AS KIND,SUBSTRING(TA003,1,6) AS MONTHS ,0 AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.ACTTA  WITH (NOLOCK),[TK].dbo.ACTTB WITH (NOLOCK)");
            FASTSQL.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
            FASTSQL.AppendFormat(@"  AND TA003 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND TB005 LIKE '4%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(TA003,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '2A' AS SEQ ,'原/物的領料' AS KIND,SUBSTRING(LA004,1,6)  AS MONTHS,SUM(LA005*LA013)*-1  AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003");
            FASTSQL.AppendFormat(@"  AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND (TE004 LIKE '1%' OR TE004 LIKE '2%')");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '3A' AS SEQ ,'鹹蛋黃' AS KIND,SUBSTRING(LA004,1,6)  AS MONTHS,SUM(LA005*LA013)*-1  AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003");
            FASTSQL.AppendFormat(@"  AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%鹹蛋黃%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '3B' AS SEQ ,'鹹蛋黃%' AS KIND, SUBSTRING(LA004,1,6)  AS MONTHS,CONVERT(DECIMAL(16,4),(SUM(LA005*LA013)*-1)/(SELECT SUM(LA005*LA013)*-1 FROM [TK].dbo.INVLA LA,[TK].dbo.MOCTE TE WHERE LA.LA006=TE.TE001 AND LA.LA007=TE.TE002 AND LA.LA008=TE.TE003 AND LA.LA004 LIKE '{0}%' AND (TE.TE004 LIKE '1%' OR TE.TE004 LIKE '2%') AND  SUBSTRING(LA.LA004,1,6)=SUBSTRING(INVLA.LA004,1,6))) *100 AS  MONEYS", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003 AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%鹹蛋黃%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '4A' AS SEQ ,'二砂' AS KIND,SUBSTRING(LA004,1,6)  AS MONTHS,SUM(LA005*LA013)*-1  AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003");
            FASTSQL.AppendFormat(@"  AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%二砂%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '4B' AS SEQ ,'二砂%' AS KIND, SUBSTRING(LA004,1,6)  AS MONTHS,CONVERT(DECIMAL(16,4),(SUM(LA005*LA013)*-1)/(SELECT SUM(LA005*LA013)*-1 FROM [TK].dbo.INVLA LA,[TK].dbo.MOCTE TE WHERE LA.LA006=TE.TE001 AND LA.LA007=TE.TE002 AND LA.LA008=TE.TE003 AND LA.LA004 LIKE '{0}%' AND (TE.TE004 LIKE '1%' OR TE.TE004 LIKE '2%') AND  SUBSTRING(LA.LA004,1,6)=SUBSTRING(INVLA.LA004,1,6))) *100 AS  MONEYS", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003 AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%二砂%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '5A' AS SEQ ,'麵粉' AS KIND,SUBSTRING(LA004,1,6)  AS MONTHS,SUM(LA005*LA013)*-1  AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003");
            FASTSQL.AppendFormat(@"  AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND (TE017 LIKE '%中筋%' OR TE017 LIKE '%低筋%'  OR TE017 LIKE '%中粉%' OR TE017 LIKE '%低粉%' OR TE017 LIKE '%強化%')");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '5B' AS SEQ ,'麵粉%' AS KIND, SUBSTRING(LA004,1,6)  AS MONTHS,CONVERT(DECIMAL(16,4),(SUM(LA005*LA013)*-1)/(SELECT SUM(LA005*LA013)*-1 FROM [TK].dbo.INVLA LA,[TK].dbo.MOCTE TE WHERE LA.LA006=TE.TE001 AND LA.LA007=TE.TE002 AND LA.LA008=TE.TE003 AND LA.LA004 LIKE '{0}%' AND (TE.TE004 LIKE '1%' OR TE.TE004 LIKE '2%') AND  SUBSTRING(LA.LA004,1,6)=SUBSTRING(INVLA.LA004,1,6))) *100 AS  MONEYS", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003 AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND (TE017 LIKE '%中筋%' OR TE017 LIKE '%低筋%'  OR TE017 LIKE '%中粉%' OR TE017 LIKE '%低粉%' OR TE017 LIKE '%強化%')");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '6A' AS SEQ ,'棕櫚油' AS KIND,SUBSTRING(LA004,1,6)  AS MONTHS,SUM(LA005*LA013)*-1  AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003");
            FASTSQL.AppendFormat(@"  AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%棕櫚油%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '6B' AS SEQ ,'棕櫚油%' AS KIND, SUBSTRING(LA004,1,6)  AS MONTHS,CONVERT(DECIMAL(16,4),(SUM(LA005*LA013)*-1)/(SELECT SUM(LA005*LA013)*-1 FROM [TK].dbo.INVLA LA,[TK].dbo.MOCTE TE WHERE LA.LA006=TE.TE001 AND LA.LA007=TE.TE002 AND LA.LA008=TE.TE003 AND LA.LA004 LIKE '{0}%' AND (TE.TE004 LIKE '1%' OR TE.TE004 LIKE '2%') AND  SUBSTRING(LA.LA004,1,6)=SUBSTRING(INVLA.LA004,1,6))) *100 AS  MONEYS", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003 AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '1%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%棕櫚油%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '7A' AS SEQ ,'袋' AS KIND,SUBSTRING(LA004,1,6)  AS MONTHS,SUM(LA005*LA013)*-1  AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE");
            FASTSQL.AppendFormat(@"  WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003");
            FASTSQL.AppendFormat(@"  AND LA004 LIKE '{0}%'",dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '2%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%袋%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  UNION ");
            FASTSQL.AppendFormat(@"  SELECT '7B' AS SEQ ,'袋%' AS KIND, SUBSTRING(LA004,1,6)  AS MONTHS,CONVERT(DECIMAL(16,4),(SUM(LA005*LA013)*-1)/(SELECT SUM(LA005*LA013)*-1 FROM [TK].dbo.INVLA LA,[TK].dbo.MOCTE TE WHERE LA.LA006=TE.TE001 AND LA.LA007=TE.TE002 AND LA.LA008=TE.TE003 AND LA.LA004 LIKE '{0}%' AND (TE.TE004 LIKE '1%' OR TE.TE004 LIKE '2%') AND  SUBSTRING(LA.LA004,1,6)=SUBSTRING(INVLA.LA004,1,6))) *100 AS  MONEYS", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  FROM [TK].dbo.INVLA,[TK].dbo.MOCTE WHERE LA006=TE001 AND LA007=TE002 AND LA008=TE003 AND LA004 LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  AND  TE004 LIKE '2%'");
            FASTSQL.AppendFormat(@"  AND TE017 LIKE '%袋%'");
            FASTSQL.AppendFormat(@"  GROUP BY SUBSTRING(LA004,1,6)");
            FASTSQL.AppendFormat(@"  ) AS TEMP ");
            FASTSQL.AppendFormat(@"  ORDER BY  SEQ,MONTHS");
            FASTSQL.AppendFormat(@"   ");
            FASTSQL.AppendFormat(@"  ");
            FASTSQL.AppendFormat(@"  ");

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
