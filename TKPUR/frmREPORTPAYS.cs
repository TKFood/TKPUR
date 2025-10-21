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
    public partial class frmREPORTPAYS : Form
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

        public frmREPORTPAYS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search( string TG002, string MA001)
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

                if (!string.IsNullOrEmpty(MA001))
                {
                    sbSqlQuery2.AppendFormat(@" 
                                            AND (廠商全名 LIKE '%{0}%' OR 供應廠商 LIKE '%{0}%')
                                                ", MA001);
                }
                else
                {
                    sbSqlQuery2.AppendFormat(@" 
                                           
                                                ");
                }

                if (!string.IsNullOrEmpty(TG002))
                {
                    sbSqlQuery3.AppendFormat(@" 
                                            AND 單號 LIKE '%{0}%'
                                                ", TG002);
                }
                else
                {
                    sbSqlQuery3.AppendFormat(@" 
                                           
                                                ");
                }
                                

                //採購的進貨+製令的託外進貨
                sbSql.AppendFormat(@"                                    
                                   SELECT *
                                    FROM 
                                    (
                                    SELECT 
                                    TG002 AS '單號'
                                    ,TG003 AS '進貨日期'                                    
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
                                    , TG001 AS '單別'
                                    ,TG005 AS '供應廠商'

                                    FROM [TK].dbo.PURTG
                                    WHERE 1=1

                                    UNION ALL
                                    SELECT 
                                    TH002 AS '單號'
                                    ,TH003 AS '進貨日期'                                    
                                    ,MA003 AS '廠商全名'
                                    ,TH014 AS '發票號碼'
                                    ,TH013 AS '發票日期'
                                    ,TH011 AS '統一編號'
                                    ,(CASE WHEN TH015=1 THEN '應稅內含' 
                                    WHEN TH015=2 THEN '應稅外加' 
                                    WHEN TH015=3 THEN '零稅率' 
                                    WHEN TH015=4 THEN '免稅' 
                                    WHEN TH015=9 THEN '不計稅' 
                                    END) AS '課稅別'
                                    ,TH031 AS '本幣貨款金額'
                                    ,TH032 AS '本幣稅額'
                                    ,(TH031+TH032) AS '本幣合計金額'
                                    , TH001 AS '單別'
                                    ,TH005 AS '供應廠商'

                                    FROM [TK].dbo.MOCTH,[TK].dbo.PURMA
                                    WHERE TH005=MA001
                                    ) AS TEMP
                                    WHERE 1=1
                                    AND 單別+單號 IN (
                                        SELECT TB005+TB006 FROM [TK].dbo.ACPTB
                                        WHERE ISNULL(TB006,'')<>''
                                    )
                                    {0}
                                    {1}
                               
                                    ", sbSqlQuery.ToString(), sbSqlQuery2.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;                  

                    MessageBox.Show("查無資料");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search( textBox5.Text.Trim(), textBox6.Text.Trim());
        }
        #endregion
    }
}
