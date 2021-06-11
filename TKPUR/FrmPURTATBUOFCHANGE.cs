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
                    

                }
                else
                {
                    textBox7.Text = "";
                    textBox8.Text = "";
                    
                }

                SearchPURTB(textBox7.Text, textBox8.Text);
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }




        #endregion

      
    }
}
