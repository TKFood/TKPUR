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
    public partial class FrmPURTATBUOFCHANGE : Form
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


        public FrmPURTATBUOFCHANGE()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT [TA001] AS '請購單',[TA002] AS '請購單號',[TA006] AS '單頭備註'
                                    FROM [TK].dbo.PURTA
                                    WHERE TA001 LIKE '%{0}%' AND TA002 LIKE '%{1}%'
                                    ORDER BY  [TA001],[TA002]
                                    ",textBox1.Text.ToString().Trim(), textBox2.Text.ToString().Trim());

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
                    textBox7.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["請購單號"].Value.ToString().Trim();

                    textBox3.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["請購單號"].Value.ToString().Trim();


                }
                else
                {
                    textBox7.Text = "";
                    textBox8.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";

                }

                SearchPURTB(textBox7.Text, textBox8.Text);

                string MAXVERSIONS = GETMAXVERSIONSPURTATBCHAGE(textBox7.Text, textBox8.Text);
                textBox9.Text = (Convert.ToInt32(MAXVERSIONS) + 1).ToString();
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
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" 
                                    SELECT 
                                    [TB003] AS '請購序號'
                                    ,[TB004] AS '品號'
                                    ,[TB005] AS '品名'
                                    ,[TB009] AS '請購教量'
                                    ,[TB011] AS '需求日'
                                    ,[TB012] AS '單身備註'
                                    FROM [TK].dbo.PURTB
                                    WHERE TB001='{0}' AND TB002='{1}'
                                    ORDER BY [TB003]
                                    ",TB001,TB002 );

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

        public string GETMAXVERSIONSPURTATBCHAGE(string TA001,string TA002)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                SqlTransaction tran;
                SqlCommand cmd = new SqlCommand();
                DataSet ds = new DataSet();
                string VERSIONS;

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;

                sqlConn = new SqlConnection(connectionString);
                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();

               
                sbSql.AppendFormat(@"  
                                    SELECT
                                    MAX([VERSIONS]) AS VERSIONS,[TA001],[TA002]
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001]='{0}' AND [TA002]='{1}'
                                    GROUP BY [TA001],[TA002]
                                    ", TA001,TA002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    VERSIONS = ds.Tables["ds"].Rows[0]["VERSIONS"].ToString();
                    return VERSIONS;

                }
                else
                {
                    return "0";
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDPURTATBCHAGE(string TA001,string TA002,string VERSIONS)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                                
                sbSql.AppendFormat(@"  
                                    INSERT INTO [TKPUR].[dbo].[PURTATBCHAGE]
                                    ([VERSIONS],[TA001],[TA002],[TA006],[TB003],[TB004],[TB005],[TB009],[TB011],[TB012])

                                    SELECT '{2}',[TA001],[TA002],[TA006],[TB003],[TB004],[TB005],[TB009],[TB011],[TB012]
                                    FROM [TK].dbo.PURTA,[TK].dbo.PURTB
                                    WHERE TA001=TB001 AND TA002=TB002
                                    AND TA001='{0}' AND TA002='{1}'
                                    ", TA001,  TA002,  VERSIONS);

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

        public void SEARCHPURTATBCHAGEVERSIONS(string TA001,string TA002)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [VERSIONS] AS '變更', [TA001] AS '請購單',[TA002] AS '請購單號'
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001] LIKE '%{0}%' AND [TA002] LIKE '%{1}%'
                                    GROUP BY [VERSIONS],[TA001],[TA002]
                                    ORDER BY [TA001],[TA002],[VERSIONS] DESC
                                    ", TA001, TA002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds.Tables["ds"];
                        dataGridView3.AutoResizeColumns();

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

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView4.DataSource = null;

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];
                    textBox10.Text = row.Cells["變更"].Value.ToString().Trim();
                    textBox11.Text = row.Cells["請購單"].Value.ToString().Trim();
                    textBox12.Text = row.Cells["請購單號"].Value.ToString().Trim();



                }
                else
                {
                    textBox7.Text = "";
                    textBox8.Text = "";
                    

                }
                SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);


            }
        }

        public void SEARCHPURTATBCHAGE(string TA001,string TA002,string VERSIONS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT [VERSIONS] AS '變更版號'
                                    ,[TA001] AS '請購單'
                                    ,[TA002] AS '請購單號'
                                    ,[TA006] AS '單頭備註'
                                    ,[TB003] AS '請購序號'
                                    ,[TB004] AS '品號'
                                    ,[TB005] AS '品名'
                                    ,[TB009] AS '請購教量'
                                    ,[TB011] AS '需求日'
                                    ,[TB012] AS '單身備註'
                                    FROM [TKPUR].[dbo].[PURTATBCHAGE]
                                    WHERE [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}'
                                    ORDER BY [TA001] ,[TA002],[TB003]
                                    ", TA001, TA002, VERSIONS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds.Tables["ds"];
                        dataGridView4.AutoResizeColumns();

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

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    textBox13.Text = row.Cells["單頭備註"].Value.ToString().Trim();
                    textBox14.Text = row.Cells["請購序號"].Value.ToString().Trim();
                    textBox15.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox16.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox17.Text = row.Cells["請購教量"].Value.ToString().Trim();
                    textBox18.Text = row.Cells["需求日"].Value.ToString().Trim();
                    textBox19.Text = row.Cells["單身備註"].Value.ToString().Trim();

                }
                else
                {
                    textBox13.Text = "";
                    textBox14.Text = "";
                    textBox15.Text = "";
                    textBox16.Text = "";
                    textBox17.Text = "";
                    textBox18.Text = "";
                    textBox19.Text = "";

                }
               

            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            textBox16.Text = SEARCHINVMB(textBox15.Text.Trim());
        }

        private void textBox21_TextChanged(object sender, EventArgs e)
        {
            textBox22.Text = SEARCHINVMB(textBox21.Text.Trim());
        }

        public string SEARCHINVMB(string MB001)
        {
            try
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

                SqlTransaction tran;
                SqlCommand cmd = new SqlCommand();
                DataSet ds = new DataSet();
                string MB002;

                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;

                sqlConn = new SqlConnection(connectionString);
                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT MB002 
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001='{0}'
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    MB002 = ds.Tables["ds"].Rows[0]["MB002"].ToString();
                    return MB002;

                }
                else
                {
                    return "";
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void UPDATEPURTATBCHAGETA006(string TA001, string TA002, string VERSIONS,string TA006)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    UPDATE  [TKPUR].[dbo].[PURTATBCHAGE]
                                    SET [TA006]='{3}'
                                    WHERE  [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}'
                                    ", TA001, TA002, VERSIONS,TA006);

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


        public void UPDATEPURTATBCHAGE(string TA001, string TA002, string VERSIONS, string TB003, string TB004, string TB005, string TB009, string TB011, string TB012)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                    UPDATE  [TKPUR].[dbo].[PURTATBCHAGE]
                                    SET [TB004]='{4}',[TB005]='{5}',[TB009]='{6}',[TB011]='{7}',[TB012]='{8}'
                                    WHERE  [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}' AND TB003='{3}'
                                   
                                    ", TA001, TA002, VERSIONS, TB003,  TB004,  TB005,  TB009,  TB011,  TB012);

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

        public void ADDPURTATBCHAGEDETAIL(string VERSIONS, string TA001, string TA002, string TA006, string TB003, string TB004, string TB005, string TB009, string TB011, string TB012)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                     INSERT INTO [TKPUR].[dbo].[PURTATBCHAGE]
                                    ([VERSIONS],[TA001],[TA002],[TA006],[TB003],[TB004],[TB005],[TB009],[TB011],[TB012])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}')
                                    ", VERSIONS, TA001, TA002, TA006, TB003, TB004, TB005, TB009, TB011, TB012);

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


        private void button4_Click(object sender, EventArgs e)
        {
            ADDPURTATBCHAGE(textBox7.Text, textBox8.Text, textBox9.Text);

            textBox9.Text = (Convert.ToInt32(textBox9.Text) + 1).ToString();

            MessageBox.Show("轉入完成");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHPURTATBCHAGEVERSIONS(textBox3.Text.Trim(), textBox4.Text.Trim());
        }



        private void button5_Click(object sender, EventArgs e)
        {
            UPDATEPURTATBCHAGETA006(textBox11.Text, textBox12.Text, textBox10.Text, textBox13.Text);

            SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);

            MessageBox.Show("完成");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UPDATEPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text, textBox14.Text, textBox15.Text, textBox16.Text, textBox17.Text, textBox18.Text, textBox19.Text);

            SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);

            MessageBox.Show("完成");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ADDPURTATBCHAGEDETAIL(textBox10.Text, textBox11.Text, textBox12.Text, textBox13.Text, textBox20.Text, textBox21.Text, textBox22.Text, textBox23.Text, textBox24.Text, textBox25.Text);
                                     
            SEARCHPURTATBCHAGE(textBox11.Text, textBox12.Text, textBox10.Text);


            MessageBox.Show("完成");
        }

        #endregion

       
    }
}
