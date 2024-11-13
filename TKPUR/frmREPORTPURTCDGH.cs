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
    public partial class frmREPORTPURTCDGH : Form
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
        int CommandTimeout = 120;

        public frmREPORTPURTCDGH()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void Search_PURTCPURTD(string TC003,string TC004,string TH004)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            StringBuilder QUERY1 = new StringBuilder();
            StringBuilder QUERY2 = new StringBuilder();
            StringBuilder QUERY3 = new StringBuilder();


            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                if(!string.IsNullOrEmpty(TC003))
                {
                    QUERY1.AppendFormat(@" AND TC003 LIKE '%{0}%'", TC003);
                }
                else
                {
                    QUERY1.AppendFormat(@"");
                }

                if (!string.IsNullOrEmpty(TC004))
                {
                    QUERY2.AppendFormat(@" AND (TC004 LIKE '%{0}%' OR MA002 LIKE '%{0}%')", TC004);
                }
                else
                {
                    QUERY2.AppendFormat(@"");
                }

                if (!string.IsNullOrEmpty(TH004))
                {
                    QUERY3.AppendFormat(@" AND  (TD004 LIKE '%{0}%' OR MB002 LIKE '%{0}%')", TH004);
                }
                else
                {
                    QUERY3.AppendFormat(@"");
                }



                sbSql.AppendFormat(@"  
                                    SELECT 
                                    TC004 AS '供應廠商'
                                    ,MA002 AS '廠商'
                                    ,TC001 AS '採購單別'
                                    ,TC002 AS '採購單號'
                                    ,TD003 AS '採購序號'
                                    ,TD004 AS '品號'
                                    ,MB002 AS '品名'
                                    ,TD008 AS '請購數量'
                                    ,TD009 AS '請購單位'
                                    ,TD012 AS '預交日'
                                    ,TD015 AS '已交數量'
                                    ,TD010 AS '請購單價'
                                    ,TD011 AS '請購金額'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA,[TK].dbo.INVMB
                                    WHERE TC001=TD001 AND TC002=TD002 
                                    AND TC004=MA001
                                    AND TD004=MB001
                                    {0}
                                    {1}
                                    {2}
                                    ORDER BY TC004,TC001,TC002,TD003

                                    ",QUERY1.ToString(), QUERY2.ToString(),QUERY3.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds.Tables["TEMPds1"];
                    dataGridView1.AutoResizeColumns();

                }
                else
                {
                    dataGridView1.DataSource = null;
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
            string TH011;
            string TH012;
            string TH013;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    TH011 = row.Cells["採購單別"].Value.ToString();
                    TH012 = row.Cells["採購單號"].Value.ToString();
                    TH013 = row.Cells["採購序號"].Value.ToString();
                    Search_PURTGPURTH(TH011, TH012, TH013);
                }
            }
        }
                
        public void Search_PURTGPURTH(string TH011,string TH012,string TH013)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"                                      
                                    SELECT 
                                    TG005 AS '供應廠商'
                                    ,MA002 AS '廠商'
                                    ,TG003 AS '進貨日期'
                                    ,TG001 AS '進貨單別'
                                    ,TG002 AS '進貨單號'
                                    ,TH003 AS '進貨序號'
                                    ,TH004 AS '品號'
                                    ,TH005 AS '品名'
                                    ,TH008 AS '單位'
                                    ,TH010 AS '批號'
                                    ,TH007 AS '進貨數量'
                                    ,TH015 AS '驗收數量'
                                    ,TH016 AS '計價數量'
                                    ,TH017 AS '驗退數量'
                                    ,TH011 AS '採購單別'
                                    ,TH012 AS '採購單號'
                                    ,TH013 AS '採購序號'
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.PURMA
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND TG005=MA001
                                    AND TH011='{0}'
                                    AND TH012='{1}'
                                    AND TH013='{2}'

                                    ORDER BY MA002,TG001,TG002,TH003

                                    ", TH011, TH012,TH013);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds.Tables["TEMPds1"];
                    dataGridView2.AutoResizeColumns();

                }
                else
                {
                    dataGridView2.DataSource = null;
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


        #endregion

        #region BUTTON
        private void button3_Click(object sender, EventArgs e)
        {
            Search_PURTCPURTD(textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim());
        }
        #endregion

       
    }
}
