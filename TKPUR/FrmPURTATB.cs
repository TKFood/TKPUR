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

                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[COMMENT] AS '備註' ");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002]='{1}'",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(@"  GROUP BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[VERSIONS],[COMMENT]");
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[VERSIONS],[COMMENT]");
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

                    textBox3.Text = row.Cells["請購單別"].Value.ToString();
                    textBox4.Text = row.Cells["請購單號"].Value.ToString();
                    textBox5.Text = row.Cells["修改次數"].Value.ToString();
                    
                    if(!string.IsNullOrEmpty(textBox3.Text)&& !string.IsNullOrEmpty(textBox4.Text) && !string.IsNullOrEmpty(textBox5.Text) )
                    {
                        Search2();
                    }
                }
                else
                {
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                  

                }
            }
        }

        public void Search2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [TA001] AS '請購單別',[TA002] AS '請購單號',[TB003] AS '序號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[NUM] AS '請購數量',CONVERT(NVARCHAR,[DATES],112) AS '日期',[VERSIONS] AS '修改次數',[ID]");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002]='{1}' AND  [VERSIONS]='{2}' ", textBox3.Text, textBox4.Text, textBox5.Text);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[VERSIONS]");
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

        public void UPDATE()
        {
            string TA001 = null;
            string TA002 = null;
            string VERSIONS = null;
            string TB003 = null;
            string MB001 = null;
            decimal NUM = 0;
            string COMMNET = textBox8.Text;

            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    TA001 = row.Cells["請購單別"].Value.ToString();
                    TA002 = row.Cells["請購單號"].Value.ToString();
                    VERSIONS = row.Cells["修改次數"].Value.ToString();
                    TB003 = row.Cells["序號"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();
                    NUM = Convert.ToDecimal(row.Cells["請購數量"].Value.ToString());

                    sbSql.AppendFormat(" UPDATE [TKPUR].[dbo].[PURTATB]");
                    sbSql.AppendFormat(" SET [NUM]='{0}' ,[COMMENT]='{1}' ", NUM, COMMNET);
                    sbSql.AppendFormat(" WHERE [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}' AND [TB003]='{3}' AND [MB001]='{4}'", TA001, TA002, VERSIONS,TB003, MB001);
                    sbSql.AppendFormat(" ");

                }

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

        public void ADD()
        {
            string DATES = null;
            string TA001 = null;
            string TA002 = null;
            string VERSIONS = null;
            string TB003 = null;
            string MB001 = null;
            string MB002 = null;
            string MB003 = null;
            string MB004 = null;
            decimal NUM = 0;
            string COMMNET = textBox8.Text;

            VERSIONS = GETMAXNO();

            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    DATES = DateTime.Now.ToString("yyyy/MM/dd");
                    TA001 = row.Cells["請購單別"].Value.ToString();
                    TA002 = row.Cells["請購單號"].Value.ToString();
                    //VERSIONS = row.Cells["修改次數"].Value.ToString();
                    TB003 = row.Cells["序號"].Value.ToString();
                    MB001 = row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();
                    MB003 = row.Cells["規格"].Value.ToString();
                    MB004 = row.Cells["單位"].Value.ToString();
                    NUM = Convert.ToDecimal(row.Cells["請購數量"].Value.ToString());

                  
                    sbSql.AppendFormat(" INSERT INTO [TKPUR].[dbo].[PURTATB]");
                    sbSql.AppendFormat(" ([DATES],[TA001],[TA002],[VERSIONS],[TB003],[MB001],[MB002],[MB003],[MB004],[NUM],[COMMENT])");
                    sbSql.AppendFormat(" VALUES");
                    sbSql.AppendFormat(" ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}')", DATES,TA001,TA002,VERSIONS, TB003,MB001, MB002,MB003,MB004,NUM, COMMNET);
                    sbSql.AppendFormat(" ");
                }

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

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX([VERSIONS]),'1') AS VERSIONS");
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

            if (VERSIONS.Equals("1"))
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
        public void SETSTATUSEDIT()
        {
            STATUS = "EDIT";
            textBoxstatus.Text = "修改中";
        }
        public void SETSTATUSADD()
        {
            STATUS = "ADD";
            textBoxstatus.Text = "新增中";
            textBox5.Text = null;
        }
        public void SETSTATUSFINALLY()
        {
            STATUS = null;
            textBoxstatus.Text = null;

            STATUS = "EDIT";
            textBoxstatus.Text = "修改中";
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

            FASTSQL.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[TB003] AS '序號',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[NUM] AS '請購數量',[ID]");
            FASTSQL.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
            FASTSQL.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002]='{1}' AND [VERSIONS]='{2}' ", RETA001, RETA002, REVERSIONS);
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
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002]='{1}'", textBox6.Text, textBox7.Text);
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();

            SETSTATUSEDIT();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SearchPURTATB();

            SETSTATUSADD();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(STATUS))
            {
                if (STATUS.Equals("EDIT"))
                {
                    UPDATE();
                }
                else if (STATUS.Equals("ADD"))
                {
                    ADD();
                }
            }
            else
            {
                MessageBox.Show("請重新查詢");
            }

            SETSTATUSFINALLY();

            MessageBox.Show("已完成");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search3();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }


        #endregion

     
    }
}
