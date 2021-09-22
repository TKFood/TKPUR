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
    public partial class FrmUPDATEPURTD : Form
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

        string TD013;
        string TD021;
        string TD023;

        public FrmUPDATEPURTD()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void Search()
        {
            ds.Clear();

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

                if(!string.IsNullOrEmpty(textBox7.Text))
                {
                    sbSqlQuery.AppendFormat(@" AND TD014 LIKE '%{0}%' ",textBox7.Text);
                }
                else
                {
                    sbSqlQuery.Clear();
                }

                sbSql.AppendFormat(@"  SELECT TD001 AS '採購單單別',TD002 AS '採購單單號',TD003 AS '採購單序號',TD014 AS '備註', TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量'");
                sbSql.AppendFormat(@"  ,TD013 AS '參考單單別',TD021 AS '參考單單號',TD023 AS '參考單序號'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTC,[TK].dbo.PURTD");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND TC003>='{0}' AND TC003<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  {0}" , sbSqlQuery.ToString());
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
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
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["採購單單別"].Value.ToString();
                    textBox2.Text = row.Cells["採購單單號"].Value.ToString();
                    textBox3.Text = row.Cells["採購單序號"].Value.ToString();
                    textBox4.Text = row.Cells["參考單單別"].Value.ToString();
                    textBox5.Text = row.Cells["參考單單號"].Value.ToString();
                    textBox6.Text = row.Cells["參考單序號"].Value.ToString();
  


                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                   

                }
            }
        }
        public void UPDATE()
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
               
                sbSql.AppendFormat(" UPDATE [TK].dbo.PURTD");
                sbSql.AppendFormat(" SET TD013='{0}',TD021='{1}',TD023='{2}'",textBox4.Text, textBox5.Text, textBox6.Text);
                sbSql.AppendFormat(" WHERE TD001='{0}' AND TD002='{1}' AND TD003='{2}'", textBox1.Text, textBox2.Text, textBox3.Text);
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            UPDATE();

            MessageBox.Show("完成");
            Search();
        }

        #endregion


    }
}
