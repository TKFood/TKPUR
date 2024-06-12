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
        }

        private void FrmTBPURCHECKFAX_Load(object sender, EventArgs e)
        {
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
                                    (CASE WHEN (SELECT COUNT(*) FROM [TKPUR].[dbo].[TBPURCHECKFAX] 
                                    WHERE REPLACE([TBPURCHECKFAX].TC001+[TBPURCHECKFAX].TC002,' ','')=REPLACE(PURTC.TC001+PURTC.TC002,' ',''))>=1 THEN 'Y' ELSE 'N' END
                                    ) AS '是否傳真'
                                    
                                    ,TC004 AS '供應廠商'
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

                    //foreach (DataGridViewRow dgRow in dataGridView1.Rows)
                    //{
                    //    //dgRow.DefaultCellStyle.ForeColor = Color.Blue;

                    //    // 确保当前行不是新行并且单元格不为空
                    //    if (!dgRow.IsNewRow && dgRow.Cells["是否傳真"].Value != null)
                    //    {
                    //        // 判断单元格的值是否为 "N"
                    //        if (dgRow.Cells["是否傳真"].Value.ToString().Trim().Equals("N"))
                    //        {
                    //            // 将这行的前景色设置成蓝色
                    //            //dgRow.DefaultCellStyle.ForeColor = Color.Blue;
                    //            dgRow.DefaultCellStyle.BackColor = Color.Pink;
                    //        }
                    //    }
                    //}
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

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            ApplyRowStyles();
        }

        private void ApplyRowStyles()
        {
            foreach (DataGridViewRow dgRow in dataGridView1.Rows)
            {
                // 确保当前行不是新行并且单元格不为空
                if (!dgRow.IsNewRow && dgRow.Cells["是否傳真"].Value != null)
                {
                    // 判断单元格的值是否为 "N"
                    if (dgRow.Cells["是否傳真"].Value.ToString().Trim().Equals("N"))
                    {
                        // 将这行的背景色设置成蓝色
                        dgRow.DefaultCellStyle.BackColor = Color.Pink;
                        dgRow.DefaultCellStyle.ForeColor = Color.Black; // 设置前景色为白色，以确保文本可见
                    }
                }
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

        public void ADD_TBPURCHECKFAX(string TC001, string TC002)
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
                                   INSERT INTO [TKPUR].[dbo].[TBPURCHECKFAX] 
                                    (
                                    TC001,
                                    TC002
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    )
                                   
                                    ", TC001,TC002 );

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

                    MessageBox.Show("完成");
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
        private void button2_Click(object sender, EventArgs e)
        {
            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox1.Text.Trim()))
            {
                ADD_TBPURCHECKFAX(textBox1.Text.Trim(), textBox2.Text.Trim());
            }

            SEARCH(dateTimePicker1.Value.ToString("yyyyMMdd"));

        }





        #endregion

       
    }
}
