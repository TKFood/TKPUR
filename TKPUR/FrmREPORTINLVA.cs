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
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
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
            FASTSQL.AppendFormat(@"  SELECT '1A' AS SEQ ,'營收' AS KIND,[YM]  AS MONTHS,[MONEYS] AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TKPUR].[dbo].[REPORTINLVA]");
            FASTSQL.AppendFormat(@"  WHERE [YM] LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '1B' AS SEQ ,'成本率' AS KIND,[YM]  AS MONTHS,[PERS] AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TKPUR].[dbo].[REPORTINLVA]");
            FASTSQL.AppendFormat(@"  WHERE [YM] LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
            FASTSQL.AppendFormat(@"  UNION");
            FASTSQL.AppendFormat(@"  SELECT '1C' AS SEQ ,'成本額' AS KIND,[YM]  AS MONTHS,[COSTS] AS  MONEYS");
            FASTSQL.AppendFormat(@"  FROM [TKPUR].[dbo].[REPORTINLVA]");
            FASTSQL.AppendFormat(@"  WHERE [YM] LIKE '{0}%'", dateTimePicker1.Value.Year.ToString());
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
            FASTSQL.AppendFormat(@"  AND (TE017 LIKE '%松高%' OR TE017 LIKE '%浤良%' OR TE017 LIKE '%宇貹%') ");
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

        public void SEARCHREPORTINLVA(string YM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [YM] AS '年月',[MONEYS] AS '營收',[PERS] AS '成本率',[COSTS] AS '成本額'");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[REPORTINLVA]");
                sbSql.AppendFormat(@"  WHERE [YM] LIKE '{0}%'",YM);
                sbSql.AppendFormat(@"  ORDER BY  [YM]");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }
                }

            }
            catch
            {

            }
            finally
            {

            }
        }
        public void ADDREPORTINLVA(string YM,string MONEYS, string PERS,string COSTS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(" INSERT INTO [TKPUR].[dbo].[REPORTINLVA]");
                sbSql.AppendFormat(" ( [YM],[MONEYS],[PERS],[COSTS])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}')",YM,MONEYS,PERS,COSTS);
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
      

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    dateTimePicker3.Value = Convert.ToDateTime(row.Cells["年月"].Value.ToString().Substring(0,4)+"/"+ row.Cells["年月"].Value.ToString().Substring(4, 2)+"/1");
                    textBox1.Text = row.Cells["營收"].Value.ToString(); 
                    textBox2.Text = row.Cells["成本率"].Value.ToString();
                    textBox3.Text = row.Cells["成本額"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                }
            }

        }

        public void  UPDATEREPORTINLVA(string YM, string MONEYS, string PERS, string COSTS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKPUR].[dbo].[REPORTINLVA]");
                sbSql.AppendFormat(" SET [MONEYS]='{0}',[PERS]='{1}',[COSTS]='{2}'",MONEYS,PERS,COSTS);
                sbSql.AppendFormat(" WHERE [YM]='{0}'",YM);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELREPORTINLVA(string YM)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" DELETE [TKPUR].[dbo].[REPORTINLVA]");
                sbSql.AppendFormat(" WHERE [YM]='{0}'",YM);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            CALPERCOSTS();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            CALPERCOSTS();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            CALPERCOSTS();
        }

        public void CALPERCOSTS()
        {
            textBox3.Text = (Convert.ToDecimal(textBox1.Text) * Convert.ToDecimal(textBox2.Text) / 100).ToString();
        }
        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHREPORTINLVA(dateTimePicker2.Value.ToString("yyyy"));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(dateTimePicker3.Value.ToString("yyyyMM"))&& !string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox3.Text))
            {
                ADDREPORTINLVA(dateTimePicker3.Value.ToString("yyyyMM"), textBox1.Text, textBox2.Text, textBox3.Text);
                SEARCHREPORTINLVA(dateTimePicker2.Value.ToString("yyyy"));
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(dateTimePicker3.Value.ToString("yyyyMM")) && !string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox3.Text))
            {
                if (!string.IsNullOrEmpty(dateTimePicker3.Value.ToString("yyyyMM")) && !string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text) && !string.IsNullOrEmpty(textBox3.Text))
                {
                    UPDATEREPORTINLVA(dateTimePicker3.Value.ToString("yyyyMM"), textBox1.Text, textBox2.Text, textBox3.Text);
                    SEARCHREPORTINLVA(dateTimePicker2.Value.ToString("yyyy"));
                }
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(dateTimePicker3.Value.ToString("yyyyMM")))
                {
                    DELREPORTINLVA(dateTimePicker3.Value.ToString("yyyyMM"));
                    SEARCHREPORTINLVA(dateTimePicker2.Value.ToString("yyyy"));
                }
                
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }



        #endregion

      
    }
}
