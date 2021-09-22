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
    public partial class FrmPURTCDMODIFY : Form
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

        public Report report1 { get; private set; }
        int result;
        Thread TD;

        string TD001;
        string TD002;
        string TD003;
        string TD012;
        string TD014;

        public FrmPURTCDMODIFY()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION

        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);


            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[KIND],[NAME] FROM [TKPUR].[dbo].[BASE] WHERE [KIND]='採購變更' ORDER BY ID   ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        //    if(!comboBox1.Text.Trim().Equals("System.Data.DataRowView"))
        //    {
        //        textBox3.Text = comboBox1.Text.Trim();
        //    }
            

            //if (string.IsNullOrEmpty(textBox3.Text))
            //{
               
            //}
        }
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

               
                sbSql.AppendFormat(@"  SELECT TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號',TD012 AS '需求日',TD014 AS '單身備註',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量',TD009 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTD");
                sbSql.AppendFormat(@"  WHERE TD001='{0}' AND TD002='{1}'",textBox1.Text,textBox2.Text);
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
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox3.Text = row.Cells["單身備註"].Value.ToString();
                    textBox4.Text = row.Cells["採購單別"].Value.ToString();
                    textBox5.Text = row.Cells["採購單號"].Value.ToString();
                    textBox6.Text = row.Cells["採購序號"].Value.ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(row.Cells["需求日"].Value.ToString().Substring(0,4)+"/"+ row.Cells["需求日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["需求日"].Value.ToString().Substring(6, 2));
                }
                else
                {
                    dataGridView1.DataSource = null;

                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;

                }
            }
        }

        public void UPDATE(string TD001,string TD002,string TD003,string TD012,string TD014)
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

                sbSql.AppendFormat(@" UPDATE [TK].dbo.PURTD ");
                sbSql.AppendFormat(@" SET TD012='{0}' ,TD014='{1}'",TD012,TD014);
                sbSql.AppendFormat(@" WHERE TD001='{0}' AND TD002='{1}' AND TD003='{2}' ",TD001,TD002,TD003);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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

        public void ADDPURTCDCHANGERECORD(string UPDATEDATES,string TD001,string TD002,string TD003,string TD014,string COMMENT)
        {
            string CHAGECOUNT= SERACHCHAGECOUNT(TD001, TD002, TD003);

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
                                INSERT INTO [TKPUR].[dbo].[PURTCDCHANGERECORD]
                                ([ID],[UPDATEDATES],[TD001],[TD002],[TD003],[TD014],[CHAGECOUNT],[COMMENT])
                                VALUES
                                ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}')
                                ",Guid.NewGuid(), UPDATEDATES, TD001, TD002, TD003, TD014, CHAGECOUNT, COMMENT);

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

        public string SERACHCHAGECOUNT(string TD001, string TD002, string TD003)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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


                sbSql.AppendFormat(@"  
                                    SELECT COUNT(ID) AS COUNTS
                                    FROM [TKPUR].[dbo].[PURTCDCHANGERECORD]
                                    WHERE [TD001]='{0}' AND [TD002]='{1}' AND [TD003]='{2}'
                                    ",TD001,TD002,TD003);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();



                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //return ds.Tables["TEMPds1"].Rows[0]["COUNTS"].ToString().Trim();

                    int counts = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["COUNTS"].ToString().Trim());
                    counts = counts + 1;
                    return counts.ToString();


                }
                else
                {
                    return "1";
                }

            }
            catch
            {
                return "0";
            }
            finally
            {

            }
        }

        public void SETFASTREPORT()
        {
            string SQL;
            string SQL2;
            report1 = new Report();
            report1.Load(@"REPORT\採購修改.frx");

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

            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();

            FASTSQL.AppendFormat(@"   
                                SELECT 
                                 [ID],CONVERT(nvarchar,[UPDATEDATES],111) AS '修改日期',[TD001] AS '採購單別',[TD002] AS '採購單號',[TD003] AS '採購序號',[TD014] AS '備註',[CHAGECOUNT] AS '修改次數',[COMMENT] AS '修改原因'
                                FROM [TKPUR].[dbo].[PURTCDCHANGERECORD]
                                WHERE CONVERT(nvarchar,[UPDATEDATES],112)>='{0}' AND  CONVERT(nvarchar,[UPDATEDATES],112)<='{1}'
                                ORDER BY [TD001],[TD002],[TD003],[CHAGECOUNT]                    
                                ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox3.Text))
            {
                UPDATE(textBox4.Text,textBox5.Text,textBox6.Text,dateTimePicker1.Value.ToString("yyyyMMdd"),textBox3.Text);
                ADDPURTCDCHANGERECORD(DateTime.Now.ToString("yyyy/MM/dd  HH:mm:ss"),textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox3.Text,comboBox1.Text.Trim());

                Search();
            }
           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }


        #endregion


    }
}
