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

        public Report report1 { get; private set; }

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
            
            if(string.IsNullOrEmpty(textBox3.Text))
            {
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
            else
            {
                
            }
            
        }


        public void SERACHCOPPURPCTPTOADDORUPDATE()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds = new DataSet();


            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"
                                    SELECT *
                                    FROM [TKPUR].[dbo].[COPPURPCT]
                                    WHERE [YM]='{0}'
                                    ", dateTimePicker4.Value.ToString("yyyyMM"));

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds.Clear();
                adapter1.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    ADDCOPPURPCT(dateTimePicker4.Value.ToString("yyyyMM"),textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim());
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        UPDATECOPPURPCT(dateTimePicker4.Value.ToString("yyyyMM"), textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim());

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

        public void ADDCOPPURPCT(string YM,string PURMONEY1,string PURMONEY2,string COPMONEY)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
               
                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKPUR].[dbo].[COPPURPCT]
                                    ([YM],[PURMONEY1],[PURMONEY2],[COPMONEY])
                                    VALUES
                                    ('{0}',{1},{2},{3})
                                    ", YM, PURMONEY1, PURMONEY2, COPMONEY);


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

        public void UPDATECOPPURPCT(string YM, string PURMONEY1, string PURMONEY2, string COPMONEY)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    UPDATE [TKPUR].[dbo].[COPPURPCT]
                                    SET [PURMONEY1]={1},[PURMONEY2]={2},[COPMONEY]={3}
                                    WHERE [YM]='{0}'
                                    ", YM, PURMONEY1, PURMONEY2, COPMONEY);


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

        public void SETFASTREPORT()
        {
            string SQL;
            string SQL2;
            report1 = new Report();
            report1.Load(@"REPORT\原物料的營收佔比.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
           
            SQL = SETFASETSQL();
            Table.SelectCommand = SQL;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();          

            FASTSQL.AppendFormat(@"   
                                SELECT 
                                [YM] AS '年月' ,[PURMONEY1] AS '原料金額',[PURMONEY2] AS '物料金額',[COPMONEY] AS '營收金額'
                                ,ROUND([PURMONEY1]/[COPMONEY],4) AS '原料佔比',ROUND([PURMONEY2]/[COPMONEY],4) AS '物料佔比'
                                FROM [TKPUR].[dbo].[COPPURPCT]
                                WHERE [YM]>='{0}' AND [YM]<='{1}'
                                ",dateTimePicker1.Value.ToString("yyyyMM"), dateTimePicker2.Value.ToString("yyyyMM"));

            return FASTSQL.ToString();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SERACHPUR();
            SERACHCOP();

            SERACHCOPPURPCTPTOADDORUPDATE();
        }

        #endregion


    }
}
