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
using FastReport.Export.Pdf;
using System.Net.Mail;
using System.Net.Mime;
using System.Diagnostics;

namespace TKPUR
{
    public partial class FrmPURTGTPURTHCHECK : Form
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

        Report report1 = new Report();


        public FrmPURTGTPURTHCHECK()
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
            Sequel.AppendFormat(@"SELECT 
                                    [ID]
                                    ,[KIND]
                                    ,[PARAID]
                                    ,[PARANAME]
                                    FROM [TKPUR].[dbo].[TBPARA]
                                    WHERE [KIND]='FrmPURTGTPURTHCHECK' ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARANAME", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARANAME";
            comboBox1.DisplayMember = "PARANAME";
            sqlConn.Close();


        }
        public void Search(string SDAY, string EDAY,string KINDS)
        {
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

                if(KINDS.Equals("未確認"))
                {
                    sbSqlQuery.AppendFormat(@" 
                                            AND REPLACE(TG001+TG002,' ','') NOT IN (
                                            SELECT
                                            REPLACE(TG001+TG002,' ','')
                                            FROM [TKPUR].[dbo].[TBPURTGCHECKS]
                                            )
                                            ");
                }
                else if (KINDS.Equals("已確認"))
                {
                    sbSqlQuery.AppendFormat(@" 
                                            AND REPLACE(TG001+TG002,' ','') IN (
                                            SELECT
                                            REPLACE(TG001+TG002,' ','')
                                            FROM [TKPUR].[dbo].[TBPURTGCHECKS]
                                            )
                                            ");
                }
                if (KINDS.Equals("全部"))
                {
                    sbSqlQuery.AppendFormat(@" 
                                           
                                            ");
                }

                sbSql.AppendFormat(@"                                    
                                    SELECT 
                                    TG001 AS '單別'
                                    ,TG002 AS '單號'
                                    ,TG003 AS '進貨日期'
                                    --,TG005 AS '供應廠商'
                                    ,TG021 AS '廠商全名'
                                    ,TG011 AS '發票號碼'
                                    ,TG027 AS '發票日期'
                                    ,TG022 AS '統一編號'
                                    ,(CASE WHEN TG010=1 THEN '應稅內含' 
                                    WHEN TG010=2 THEN '應稅外加' 
                                    WHEN TG010=3 THEN '零稅率' 
                                    WHEN TG010=4 THEN '免稅' 
                                    WHEN TG010=9 THEN '不計稅' 
                                    END) AS '課稅別'
                                    ,TG031 AS '本幣貨款金額'
                                    ,TG032 AS '本幣稅額'
                                    ,(TG031+TG032) AS '本幣合計金額'
                                    
                                    FROM [TK].dbo.PURTG
                                    WHERE TG003>='{0}' AND TG003<='{1}'
                                    {2}

                                    ", SDAY, EDAY, sbSqlQuery.ToString());

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

                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        // 設定數字格式
                        // 或使用 "N2" 表示兩位小數點（例如：12,345.67）
                        dataGridView1.Columns["本幣貨款金額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點
                        dataGridView1.Columns["本幣稅額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點
                        dataGridView1.Columns["本幣合計金額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點



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
            textBox1.Text = null;
            textBox2.Text = null;
           

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["單別"].Value.ToString();
                    textBox2.Text = row.Cells["單號"].Value.ToString();

                    SEARCH_PURTH(row.Cells["單別"].Value.ToString(), row.Cells["單號"].Value.ToString());
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;                   
                }
            }
        }

        public void SEARCH_PURTH(string TH001,string TH002)
        {
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
                                    TH003 AS '序號'
                                    ,TH004 AS '品號'
                                    ,TH005 AS '品名'
                                    ,TH006 AS '規格'
                                    ,TH007 AS '進貨數量'
                                    ,TH016 AS '計價數量'
                                    ,TH008 AS '單位'
                                    ,TH010 AS '批號'
                                    ,TH018 AS '原幣單位進價'
                                    ,TH047 AS '本幣未稅金額'
                                    ,TH048 AS '本幣稅額'
                                    FROM [TK].dbo.PURTH
                                    WHERE TH001='{0}' AND TH002='{1}'


                                    ", TH001, TH002);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["TEMPds1"];
                        dataGridView2.AutoResizeColumns();

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        // 設定數字格式
                        // 或使用 "N2" 表示兩位小數點（例如：12,345.67）
                        dataGridView2.Columns["進貨數量"].DefaultCellStyle.Format = "N2"; // 每三位一個逗號，無小數點
                        dataGridView2.Columns["計價數量"].DefaultCellStyle.Format = "N2"; // 每三位一個逗號，無小數點
                        dataGridView2.Columns["原幣單位進價"].DefaultCellStyle.Format = "N2"; // 每三位一個逗號，無小數點
                        dataGridView2.Columns["本幣未稅金額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點
                        dataGridView2.Columns["本幣稅額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點



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

        public void ADD_CHECK_PURTG(string TG001,string TG002)
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
                                    INSERT INTO [TKPUR].[dbo].[TBPURTGCHECKS]
                                    ([TG001]
                                    ,[TG002])
                                    VALUES
                                    ('{0}','{1}')
                                    ", TG001, TG002);


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELETE_CHECK_PURTG(string TG001, string TG002)
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
                                    DELETE [TKPUR].[dbo].[TBPURTGCHECKS]
                                    WHERE TG001='{0}' AND TG002='{1}'
                                    ", TG001, TG002);


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
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
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"),comboBox1.Text.ToString());
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //新增確認單號
            ADD_CHECK_PURTG(textBox1.Text.Trim(), textBox2.Text.Trim());

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            //MessageBox.Show("完成");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //解除確認單號
            DELETE_CHECK_PURTG(textBox1.Text.Trim(), textBox2.Text.Trim());

            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
            //MessageBox.Show("完成");
        }
        #endregion


    }
}
