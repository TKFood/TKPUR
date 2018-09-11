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
    public partial class FrmCALINV : Form
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

        public FrmCALINV()
        {
            InitializeComponent();

            if (!string.IsNullOrEmpty(comboBox1.Text))
            {
                label2.Text = searchMB002(comboBox1.Text);

                if (comboBox1.Text.Equals("40100113000016"))
                {
                    textBox2.Text = "203021091";

                    label6.Text = searchMB002(textBox2.Text);
                }
            }
        }

        #region FUNCTION
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(comboBox1.Text))
            {
                label2.Text=searchMB002(comboBox1.Text);

                if(comboBox1.Text.Equals("40100113000016"))
                {
                    textBox2.Text = "203021091";

                    label6.Text = searchMB002(textBox2.Text);
                }
            }

           
        }

        public string searchMB002(string MB001)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            sbSql.Clear();
            sbSqlQuery.Clear();

           
            sbSql.AppendFormat(@" SELECT MB001,MB002 FROM [TK].dbo.INVMB WHERE MB001='{0}' ", MB001);

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

            sqlCmdBuilder = new SqlCommandBuilder(adapter);
            sqlConn.Open();
            ds.Clear();
            adapter.Fill(ds, "TEMPds1");
            sqlConn.Close();

            if (ds.Tables["TEMPds1"].Rows.Count == 0)
            {
                return null;
            }
            else
            {
                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    foreach (DataRow od2 in ds.Tables["TEMPds1"].Rows)
                    {                       
                        return od2["MB002"].ToString();                    
                    }
                }
                return null;
            }

        }
        
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {

        }
        #endregion

        
    }
}
