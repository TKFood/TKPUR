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


namespace TKPUR
{
    public partial class FrmPURTATBTB003 : Form
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


        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;
        public Report report1 { get; private set; }
        string REPORTID;
        string DELID;

        int ROWSINDEX = 0;
        int COLUMNSINDEX = 0;


        public FrmPURTATBTB003()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search(string TA001,string TA002,string TA007)
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

                if(!string.IsNullOrEmpty(TA001))
                {
                    sbSqlQuery.AppendFormat(@" AND TA001 LIKE '%{0}%'", TA001);
                }
                else
                {
                    sbSqlQuery.AppendFormat(@" ");
                }

                if(!string.IsNullOrEmpty(TA002))
                {
                    sbSqlQuery.AppendFormat(@" AND TA002 LIKE '%{0}%'", TA002);
                }
                else
                {
                    sbSqlQuery.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(TA007))
                {
                    sbSqlQuery.AppendFormat(@" AND TA007='N'");
                }
                else
                {
                    sbSqlQuery.AppendFormat(@" ");
                }


                sbSql.AppendFormat(@"

                                  SELECT TA001 AS '請購單別',TA002 AS '請購單號'
                                    FROM [TK].dbo.PURTA
                                    WHERE 1=1
                                    {0}

                                    ORDER BY TA001,TA002
                                    ", sbSqlQuery.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox3.Text = row.Cells["請購單別"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["請購單號"].Value.ToString().Trim();

                }
                else
                {
                    textBox3.Text = "";
                    textBox4.Text = "";
                  

                }

                SearchPURTB(textBox3.Text, textBox4.Text);

            }
        }

        public void SearchPURTB(string TB001,string TB002)
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
                                    SELECT TB003 AS '序號',TB004 AS '品號',TB005 AS '品名',TB007 AS '請購單位',TB009 AS '請購數量',TB011 AS '需求日期',MA002 AS '供應廠商',TB010 
                                    FROM [TK].dbo.PURTB,[TK].dbo.PURMA
                                    WHERE TB010=MA001
                                    AND TB001='{0}' AND TB002='{1}'
                                    ", TB001,TB002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["ds"];
                        dataGridView2.AutoResizeColumns();

                        dataGridView2.Columns["序號"].Width = 50;
                        dataGridView2.Columns["請購單位"].Width = 100;
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

        public void UPDATEPURTBTB003(string TB001,string TB002)
        {
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                   ;WITH CTE AS (
	                                    SELECT TB001,TB002,TB003,TB010,TB011,RIGHT('0000'+CAST(ROW_NUMBER() OVER(ORDER BY TB010,TB011,TB003) AS nvarchar(50)),4) AS NEWTB003
	                                    FROM [TK].dbo.PURTB
	                                    WHERE TB025='N'
	                                    AND TB001='{0}' AND TB002='{1}'	
                                    )
                                    UPDATE [TK].dbo.PURTB
                                    SET PURTB.TB042=PURTB.TB003,PURTB.TB003=CTE.NEWTB003
                                    FROM CTE,[TK].dbo.PURTB
                                    WHERE CTE.TB001=PURTB.TB001 AND CTE.TB002=PURTB.TB002  AND CTE.TB003=PURTB.TB003
                                    AND PURTB.TB025='N'
                                    AND PURTB.TB001='{0}' AND  PURTB.TB002='{1}'
                                 
                                    ",TB001,TB002);

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
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox1.Text,textBox2.Text,"N");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            UPDATEPURTBTB003(textBox3.Text, textBox4.Text);
            SearchPURTB(textBox3.Text, textBox4.Text);
        }
        #endregion


    }
}
