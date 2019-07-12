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

            
        }

        #region FUNCTION
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
           
        }
        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            label2.Text=searchMB002(textBox2.Text);
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
        
        public void CALINV()
        {
            int NUM = Convert.ToInt32(textBox1.Text);

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  WITH TEMPTABLE (MD001,MD003,MD004,MD006,MD007,MD008,MC004,NUM,LV) AS");
                sbSql.AppendFormat(@"  (");
                sbSql.AppendFormat(@"  SELECT  MD001,MD003,MD004,MD006,MD007,MD008,MC004,CONVERT(decimal(18,5),(MD006*(1+MD008)/MD007)/MC004)*{1} AS NUM,1 AS LV FROM [TK].dbo.VBOMMD WHERE  MD001='{0}'", textBox2.Text, NUM);
                sbSql.AppendFormat(@"  UNION ALL");
                sbSql.AppendFormat(@"  SELECT A.MD001,A.MD003,A.MD004,A.MD006,A.MD007,A.MD008,A.MC004,CONVERT(decimal(18,5),(A.MD006*(1+A.MD008)/A.MD007/A.MC004)*(B.NUM))*{0} AS NUM,LV+1", NUM);
                sbSql.AppendFormat(@"  FROM [TK].dbo.VBOMMD A");
                sbSql.AppendFormat(@"  INNER JOIN TEMPTABLE B on A.MD001=B.MD003");
                sbSql.AppendFormat(@"  )");
                sbSql.AppendFormat(@"    SELECT MD003 AS '物料',MB002 AS '品名',MD004 AS '單位',NUM AS '需求量', MD001, MD006,MD007,MD008,MC004,LV ");
                sbSql.AppendFormat(@"  FROM TEMPTABLE ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.INVMB ON MB001=MD003");
                sbSql.AppendFormat(@"  WHERE  (MD003 LIKE '2%') ");
                //sbSql.AppendFormat(@"  WHERE  (MD003 LIKE '{0}%') ", textBox2.Text);
                sbSql.AppendFormat(@"  ORDER BY LV,MD001,MD003");



                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "TEMPds2");
                sqlConn.Close();

              
                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView1.AutoResizeColumns();
                        //dataGridView1.CurrentCell = dataGridView1[0, rownum];

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
            CALINV();
        }

        #endregion

       
    }
}
