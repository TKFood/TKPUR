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
    public partial class FrmPURTATB : Form
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
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5= new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5= new DataSet();
        DataTable dt = new DataTable();
        DataTable dtADD = new DataTable();

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;
        public Report report1 { get; private set; }
        string RETA001;
        string RETA002;
        string REVERSIONS;
        string DELID;

        public FrmPURTATB()
        {
            InitializeComponent();
            SETdtADD();
        }

        #region FUNCTION
        public void SETdtADD()
        {
            dtADD.Columns.Add("日期", typeof(String));
            dtADD.Columns.Add("請購單別", typeof(String));
            dtADD.Columns.Add("請購單號", typeof(String));
            dtADD.Columns.Add("修改次數", typeof(String));
            dtADD.Columns.Add("品號", typeof(String));
            dtADD.Columns.Add("品名", typeof(String));
            dtADD.Columns.Add("規格", typeof(String));
            dtADD.Columns.Add("單位", typeof(String));
            dtADD.Columns.Add("請購數量", typeof(String));

        }
        public void Search()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[COMMENT] AS '單頭備註', [ID] ");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002] LIKE '%{1}%' ",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[COMMENT]");
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
            dataGridView2.DataSource = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox9.Text = row.Cells["請購單別"].Value.ToString();
                    textBox10.Text = row.Cells["請購單號"].Value.ToString();
                    textBox11.Text = row.Cells["單頭備註"].Value.ToString();
                    textBox12.Text = row.Cells["ID"].Value.ToString();
                    DELID = row.Cells["ID"].Value.ToString();
                  

                    if (!string.IsNullOrEmpty(row.Cells["ID"].Value.ToString()) )
                    {
                        Search2(row.Cells["ID"].Value.ToString());
                    }
                }
                else
                {
                    dataGridView2.DataSource = null;

                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBox12.Text = null;
                    DELID = null;



                }
            }
        }

        public void Search2(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [TA001] AS '請購單別',[TA002] AS '請購單號',[TA003] AS '序號',[COMMENTD] AS '單身備註',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[NUM] AS '請購數量',CONVERT(NVARCHAR,[DATES],112) AS '日期',[ID],[MID]");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat(@"  WHERE [MID]='{0}' ", ID);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[ID]");
                sbSql.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["ds2"];
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

        public void UPDATEPURTATB(string ID,string COMMENT)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(" SET [COMMENT]='{0}' ", COMMENT);
                sbSql.AppendFormat(" WHERE[ID] ='{0}'", ID);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat("  ");

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

        public void ADDPURTATB(string TA001,string TA002,string COMMENT)
        {
            Guid Guid = Guid.NewGuid();

            try
            {            
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(" ([ID],[DATES],[TA001],[TA002],[COMMENT])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}',getdate(),'{1}','{2}','{3}')",Guid.ToString() , TA001, TA002, COMMENT);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat(" ([ID],[MID],[DATES],[TA001],[TA002],[TA003],[MB001],[MB002],[MB003],[MB004],[NUM],[COMMENTD])");
                sbSql.AppendFormat(" SELECT NEWID(),'{0}',GETDATE(),TB001,TB002,TB003,TB004,TB005,TB006,TB007,TB009,TB012", Guid.ToString());
                sbSql.AppendFormat(" FROM [TK].dbo.PURTB");
                sbSql.AppendFormat(" WHERE TB001='{0}' AND TB002 ='{1}'",TA001,TA002);
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

        public string GETMAXNO()
        {
            string VERSIONS;
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX([VERSIONS]),'0') AS VERSIONS");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA002='{1}'", textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        VERSIONS = SETVERSION(ds4.Tables["TEMPds4"].Rows[0]["VERSIONS"].ToString());
                        return VERSIONS;

                    }
                    return null;
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

        public string SETVERSION(string VERSIONS)
        {

            if (VERSIONS.Equals("0"))
            {
                return "1";
            }

            else
            {
                int serno = Convert.ToInt16(VERSIONS);
                serno = serno + 1;
              
                return serno.ToString();
            }
        }
       
        public void SETSTATUSFINALLY()
        {
            STATUS = null;
            

            STATUS = "EDIT";

        }

        public void SearchPURTATB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TB001 AS '請購單別',TB002 AS '請購單號',TB003 AS '序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB007 AS '單位',TB009 AS '請購數量' , CONVERT(NVARCHAR,TA003,112) AS '日期' ,'' AS '修改次數'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTB, [TK].dbo.PURTA");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TB001='{0}' AND TB002='{1}' ",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(@"  ORDER BY  CONVERT(NVARCHAR,TA003,112) ,TB001,TB002");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds3.Tables["ds3"];
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

        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\請購變更單.frx");

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

            FASTSQL.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[TB003] AS '序號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[NUM] AS '請購數量',[ID],[COMMENT]  AS '備註' ");
            FASTSQL.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
            FASTSQL.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002] LIKE '%{1}%' AND [VERSIONS]='{2}' ", RETA001, RETA002, REVERSIONS);
            FASTSQL.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[VERSIONS],[TB003] ");
            FASTSQL.AppendFormat(@"   ");

            return FASTSQL.ToString();
        }

        public void Search3()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數'");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002] LIKE '%{1}%'", textBox6.Text, textBox7.Text);
                sbSql.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[VERSIONS]");
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[VERSIONS]");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds4.Tables["ds4"];
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

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox3.Text = row.Cells["請購單別"].Value.ToString();
                    textBox4.Text = row.Cells["請購單號"].Value.ToString();
                    textBox5.Text = row.Cells["序號"].Value.ToString();
                    textBox8.Text = row.Cells["單身備註"].Value.ToString();
                    textBox13.Text = row.Cells["ID"].Value.ToString();



                }
                else
                {
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox8.Text = null;
                    textBox13.Text = null;

                }
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
           
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    RETA001 = row.Cells["請購單別"].Value.ToString();
                    RETA002 = row.Cells["請購單號"].Value.ToString();
                    REVERSIONS = row.Cells["修改次數"].Value.ToString();

                   
                }
                else
                {
                    RETA001 = null;
                    RETA002 = null;
                    REVERSIONS = null;

                }
            }
        }

        public void DELPURTATB(string ID)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  DELETE [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat("  WHERE [ID]='{0}' ", ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  DELETE [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat("  WHERE [MID]='{0}' ", ID);
                sbSql.AppendFormat("  ");

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

        public void UPDATEPURTATBD(string ID,string COMMENT)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  UPDATE [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat("  SET [COMMENTD]='{0}'",COMMENT);
                sbSql.AppendFormat("  WHERE [ID]='{0}'",ID);
                sbSql.AppendFormat("  ");

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
            if(!string.IsNullOrEmpty(textBox9.Text)&& !string.IsNullOrEmpty(textBox10.Text)&& !string.IsNullOrEmpty(textBox11.Text))
            {
                ADDPURTATB(textBox9.Text, textBox10.Text, textBox11.Text);
                Search();
            }
           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox13.Text) && !string.IsNullOrEmpty(textBox8.Text))
            {
                UPDATEPURTATBD(textBox13.Text, textBox8.Text);
                Search();
            }


        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search3();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrEmpty(textBox11.Text) )
            {
                UPDATEPURTATB(textBox12.Text, textBox11.Text);
                Search();
            }
               
        }
        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if(!string.IsNullOrEmpty(textBox12.Text))
                {
                    DELPURTATB(textBox12.Text);

                    Search();
                }               

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }


        #endregion

      
    }
}
