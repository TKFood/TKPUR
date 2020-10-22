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
    public partial class FrmREOIRTCOPPURPCT : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        DataSet ds = new DataSet();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public FrmREOIRTCOPPURPCT()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now;
        }

        public void SERACHPUR()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds = new DataSet();


            ds.Clear();
            textBox1.Text = "0";
            textBox2.Text = "0";

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                
                sbSql.AppendFormat(@"  
                                    SELECT SUBSTRING(TH004,1,1) AS 'KINDS',CONVERT(INT,SUM(TH047+TH048)) AS 'MONEYS'
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND SUBSTRING(TG003,1,6)='{0}' 
                                    AND (TH004 LIKE '1%' OR TH004 LIKE '2%' )
                                    GROUP BY SUBSTRING(TH004,1,1)"
                                    , dateTimePicker4.Value.ToString("yyyyMM"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                        textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["MONEYS"].ToString();
                        textBox2.Text = ds.Tables["TEMPds1"].Rows[1]["MONEYS"].ToString();
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

        public void SERACHCOP()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds = new DataSet();


            ds.Clear();
            textBox3.Text = "0";
            

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"
                                    SELECT CONVERT(INT,SUM(MONEYS)) MONEYS
                                    FROM(
                                    SELECT SUM(TH037) AS 'MONEYS'
                                    FROM[TK].dbo.COPTG,[TK].dbo.COPTH
                                    WHERE TG001 = TH001 AND TG002 = TH002
                                    AND  SUBSTRING(TG003, 1, 6) = '{0}'
                                    UNION ALL
                                    SELECT SUM(TB031) AS 'MONEYS'
                                    FROM[TK].dbo.POSTB
                                    WHERE  SUBSTRING(TB001, 1, 6) = '{0}'
                                    ) AS TEMP"
                                    , dateTimePicker4.Value.ToString("yyyyMM"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {

                        textBox3.Text = ds.Tables["TEMPds1"].Rows[0]["MONEYS"].ToString();
                       
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

        }

        private void button4_Click(object sender, EventArgs e)
        {
            SERACHPUR();
            SERACHCOP();        }

        #endregion


    }
}
