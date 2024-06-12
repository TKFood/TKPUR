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
using System.Xml;
using TKITDLL;
using System.Globalization;

namespace TKPUR
{
    public partial class FrmTBPURCHECKFAX : Form
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
        int result;

        public FrmTBPURCHECKFAX()
        {
            InitializeComponent();

            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }


        #region FUNCTION
        public void SEARCH(string SDATES)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
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
                                    TC004 AS '供應廠商'
                                    ,MA002 AS '廠商'
                                    ,TC001 AS '採購單別'
                                    ,TC002 AS '採購單號'

                                    FROM [TK].dbo.PURTC,[TK].dbo.PURMA
                                    WHERE 1=1
                                    AND MA001=TC004
                                    AND TC002 LIKE '{0}%'
                                    ORDER BY TC001,TC002

                                    ", SDATES);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds.Tables["ds"];
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

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["採購單別"].Value.ToString().Trim();
                    textBox2.Text = row.Cells["採購單號"].Value.ToString().Trim();

                    SEARCH_PURTC_PURTD(textBox1.Text.Trim(), textBox2.Text.Trim());

                }
                else
                {
                    textBox1.Text = "";
                    textBox2.Text = "";
                    dataGridView2.DataSource = null;
                }


            }
        }

        public void SEARCH_PURTC_PURTD(string TC001,string TC002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
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
                                    TD012 AS '預交日'                                
                                    ,TD001 AS '採購單別'
                                    ,TD002 AS '採購單號'
                                    ,TD003 AS '序號'
                                    ,TD004 AS '品號'
                                    ,TD005 AS '品名'
                                    ,TD006 AS '規格'
                                    ,TD008 AS '採購數量'
                                    ,TD009 AS '單位'
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.PURMA
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MA001=TC004
                                    AND TC014='Y'
                                    AND TC001='{0}' AND TC002='{1}'
                                    ORDER BY TD001,TD002,TD003

                                    ", TC001, TC002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds.Tables["ds"];
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

            }
        }

        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        private void button1_Click(object sender, EventArgs e)
        {

        }



        #endregion

        
    }
}
