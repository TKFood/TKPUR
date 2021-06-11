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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }




        #endregion
    }
}
