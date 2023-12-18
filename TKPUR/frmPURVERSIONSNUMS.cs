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
    public partial class frmPURVERSIONSNUMS : Form
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

        public frmPURVERSIONSNUMS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
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
            Sequel.AppendFormat(@"
                                SELECT  [ID],[KIND],[PARAID],[PARANAME] FROM [TKPUR].[dbo].[TBPARA] WHERE [KIND]='是否結案' ORDER BY ID
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARANAME";
            comboBox1.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void comboBox2load()
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
            Sequel.AppendFormat(@"
                                SELECT  [ID],[KIND],[PARAID],[PARANAME] FROM [TKPUR].[dbo].[TBPARA] WHERE [KIND]='是否結案2' ORDER BY ID
                                ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "PARANAME";
            comboBox2.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void SEARCH_PURVERSIONSNUMS(string NAMES, string MB001,string ISCLOSE)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();
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

                if(!string.IsNullOrEmpty(NAMES))
                {
                    sbSqlQuery1.AppendFormat(@" AND [NAMES] LIKE '%{0}%'", NAMES);
                }
                else
                {
                    sbSqlQuery1.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(MB001))
                {
                    sbSqlQuery2.AppendFormat(@" AND [MB001] LIKE '%{0}%'", MB001);
                }
                else
                {
                    sbSqlQuery2.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(ISCLOSE)&& ISCLOSE.Equals("全部"))
                {
                    sbSqlQuery3.AppendFormat(@"" );
                }
                else if (!string.IsNullOrEmpty(ISCLOSE) )
                { 
                    sbSqlQuery3.AppendFormat(@" AND [ISCLOSE] LIKE '%{0}%'", ISCLOSE);
                }

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"
                                    SELECT 
                                     [NAMES] AS '版型' 
                                    ,[MB001] AS '品號' 
                                    ,[MB002] AS '品名' 
                                    ,[BACKMONEYS] AS '可退還的版費' 
                                    ,[TARGETNUMS] AS '目標進貨量' 
                                    ,[TOTALNUMS] AS '已進貨量' 
                                    ,[ISCLOSE] AS '是否結案' 
                                    FROM [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    {2}
                                    ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString());

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

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];   
                    textBox3.Text = row.Cells["版型"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["品號"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["品名"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["可退還的版費"].Value.ToString().Trim();
                    textBox7.Text = row.Cells["目標進貨量"].Value.ToString().Trim();
                    textBox8.Text = row.Cells["已進貨量"].Value.ToString().Trim();
                    comboBox2.Text = row.Cells["是否結案"].Value.ToString().Trim();


                }
                else
                {
                    textBox3.Text = "";
                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";
                    textBox7.Text = "";
                    textBox8.Text = "";

                }

               
            }
        }

        public void ADD_PURVERSIONSNUMS(string NAMES, string MB001, string MB002, string BACKMONEYS, string TARGETNUMS,string TOTALNUMS, string ISCLOSE)
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
                                   INSERT INTO [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    ([NAMES],[MB001],[MB002],[BACKMONEYS],[TARGETNUMS],[TOTALNUMS],[ISCLOSE])
                                    VALUES
                                    ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')
                                    ", NAMES, MB001, MB002, BACKMONEYS, TARGETNUMS, TOTALNUMS, ISCLOSE);

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
        public void UPDATE_PURVERSIONSNUMS(string NAMES, string MB001, string MB002, string BACKMONEYS, string TARGETNUMS, string TOTALNUMS, string ISCLOSE)
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
                                   
                                    UPDATE  [TKPUR].[dbo].[PURVERSIONSNUMS]
                                    SET MB001='{1}',MB002='{2}',BACKMONEYS='{3}',TARGETNUMS='{4}',TOTALNUMS='{5}',ISCLOSE='{6}'
                                    WHERE NAMES='{0}'
                                    ", NAMES, MB001, MB002, BACKMONEYS, TARGETNUMS, TOTALNUMS, ISCLOSE);

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
        public void DELETE_PURVERSIONSNUMS(string NAMES)
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
                                    DELETE  [TKPUR].[dbo].[PURVERSIONSNUMS]                                    
                                    WHERE NAMES='{0}'
                                    ", NAMES);

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
            SEARCH_PURVERSIONSNUMS(textBox1.Text.Trim(), textBox2.Text.Trim(),comboBox1.Text.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ADD_PURVERSIONSNUMS(textBox3.Text.Trim(), textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim(), textBox8.Text.Trim(),comboBox2.Text.ToString());
            SEARCH_PURVERSIONSNUMS(textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString());

        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATE_PURVERSIONSNUMS(textBox3.Text.Trim(), textBox4.Text.Trim(), textBox5.Text.Trim(), textBox6.Text.Trim(), textBox7.Text.Trim(), textBox8.Text.Trim(), comboBox2.Text.ToString());
            SEARCH_PURVERSIONSNUMS(textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString());

        }

        private void button4_Click(object sender, EventArgs e)
        {
            DELETE_PURVERSIONSNUMS(textBox3.Text.Trim());
            SEARCH_PURVERSIONSNUMS(textBox1.Text.Trim(), textBox2.Text.Trim(), comboBox1.Text.ToString());

        }


        #endregion


    }
}
