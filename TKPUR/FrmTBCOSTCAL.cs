using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKPUR
{
    public partial class FrmTBCOSTCAL : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        SqlTransaction tran;       
        int result;

        public FrmTBCOSTCAL()
        {
            InitializeComponent();
        }

        #region FUNCTION

        private void FrmTBCOSTCAL_Load(object sender, EventArgs e)
        {
            SEARCHTBCOSTCAL();
        }

        public void SEARCHTBCOSTCAL()
        {
            SqlConnection sqlConn = new SqlConnection();
            string connectionString;
            StringBuilder sbSql = new StringBuilder();

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);


                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT 
                                    [TYPES] AS '類別'
                                    ,[PRODPCTS]  AS '佔製造成本百分比'
                                    ,[COMPCTS] AS '佔營業成本百分比 A'
                                    ,[ITEMS] AS '細項目'
                                    ,[INCOMEPCTS] AS '進貨金額佔類別平均% B'
                                    ,[ADDPCTS] AS '調幅增加(減少)% C'
                                    ,[TPCTS] AS '影響成本率增加(減少)% D=A*B*C'
                                    ,[ID]
                                    ,[SORTS]
                                    FROM [TKPUR].[dbo].[TBCOSTCAL]
                                    ORDER BY [SORTS]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds"];
                        dataGridView1.AutoResizeColumns();


                        dataGridView1.AutoResizeColumns();
                        dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        dataGridView1.Columns["類別"].Width = 100;
                        dataGridView1.Columns["細項目"].Width = 200;

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
            if(dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox0.Text = row.Cells["ID"].Value.ToString();
                    textBox1.Text = row.Cells["類別"].Value.ToString();
                    textBox2.Text = row.Cells["佔製造成本百分比"].Value.ToString();
                    textBox3.Text = row.Cells["佔營業成本百分比 A"].Value.ToString();
                    textBox4.Text = row.Cells["細項目"].Value.ToString();
                    textBox5.Text = row.Cells["進貨金額佔類別平均% B"].Value.ToString();
                    textBox6.Text = row.Cells["調幅增加(減少)% C"].Value.ToString();
                    textBox7.Text = row.Cells["影響成本率增加(減少)% D=A*B*C"].Value.ToString();


                }
                else
                {
                    textBox0.Text = null;
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                }
            }
        }

        public void UPDATETBCOSTCAL(string ID,decimal ADDPCTS)
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


                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

              
                sbSql.AppendFormat(@" 
                                    UPDATE [TKPUR].[dbo].[TBCOSTCAL]
                                    SET [ADDPCTS]={1}
                                    WHERE [ID]='{0}'

                                    UPDATE [TKPUR].[dbo].[TBCOSTCAL]
                                    SET [TPCTS]=[COMPCTS]*[INCOMEPCTS]*[ADDPCTS]/10000
                                    WHERE [ID]='{0}'
                                    ",ID, ADDPCTS);

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
            SEARCHTBCOSTCAL();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            UPDATETBCOSTCAL(textBox0.Text.Trim(),Convert.ToDecimal(textBox6.Text.Trim()));

            SEARCHTBCOSTCAL();
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        #endregion

       
    }
}
