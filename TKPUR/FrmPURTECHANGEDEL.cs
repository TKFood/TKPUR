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
    public partial class FrmPURTECHANGEDEL : Form
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

        public FrmPURTECHANGEDEL()
        {
            InitializeComponent();
        }
        private void FrmPURTECHANGEDEL_Load(object sender, EventArgs e)
        {
            SETDATES();
        }
        #region FUNCTION
        public void SETDATES()
        {
            textBox1.Text = DateTime.Now.ToString("yyyy");
        }
        public void Search(string YEARS)
        {
            DataSet ds = new DataSet();

            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();

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
                                    TE001 AS '採購變更單別'
                                    ,TE002 AS '採購變更單號'
                                    ,TE003 AS '版次'
                                    ,[View_TB_WKF_TASK_PURTACHANGE].[TA001] AS '請購單別'
                                    ,[View_TB_WKF_TASK_PURTACHANGE].[TA002] AS '請購單號'
                                    ,[View_TB_WKF_TASK_PURTACHANGE].[VERSIONS] AS '請購版次'
                                    ,[View_TB_WKF_TASK_PURTACHANGE].[DOC_NBR] AS '請購表單號'
                                    
                                    FROM [TK].dbo.PURTE,[TK].dbo.PURTF
                                    LEFT JOIN [192.168.1.223].[UOF].[dbo].[View_TB_WKF_TASK_PURTACHANGE] ON [View_TB_WKF_TASK_PURTACHANGE].[VERSIONS]+[View_TB_WKF_TASK_PURTACHANGE].[TA001]+[View_TB_WKF_TASK_PURTACHANGE].[TA002]=SUBSTRING(PURTF.UDF01,0,17) COLLATE Chinese_Taiwan_Stroke_BIN
                                    WHERE TE001=TF001 AND TE002=TF002
                                    AND TE017='N'
                                    AND TE002 LIKE '%{0}%'
                                    ORDER BY TE001,TE002,SUBSTRING(PURTF.UDF01,0,17)
                                    ", YEARS);

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
            textBox2.Text = null;
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;
            textBox7.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox2.Text = row.Cells["採購變更單別"].Value.ToString();
                    textBox3.Text = row.Cells["採購變更單號"].Value.ToString();
                    textBox4.Text = row.Cells["版次"].Value.ToString();
                    textBox5.Text = row.Cells["請購單別"].Value.ToString();
                    textBox6.Text = row.Cells["請購單號"].Value.ToString();
                    textBox7.Text = row.Cells["請購版次"].Value.ToString();

                }
                else
                {
                    textBox2.Text = null;
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;
                    textBox7.Text = null;
                }
            }
        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox1.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}
