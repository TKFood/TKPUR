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

namespace TKPUR
{
    public partial class FrmPURSEND : Form
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

        public FrmPURSEND()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {

            string SQL;
            report1 = new Report();
            report1.Load(@"REPORT\採購單憑証.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            SQL = SETFASETSQL();

            Table.SelectCommand = SQL;

            //report1.Preview = previewControl1;
            //report1.Show();

            //// prepare a report
            //report1.Prepare();
            //// create an instance of HTML export filter
            //FastReport.Export.Pdf.PDFExport PDFEXPORT = new FastReport.Export.Pdf.PDFExport();
            //// show the export options dialog and do the export
            //if (PDFEXPORT.ShowDialog())
            //{
            //    report1.Export(PDFEXPORT, "PDFEXPORT.pdf");
            //}


            report1.PrintSettings.ShowDialog = false;
            report1.Prepare();    // show progress dialog
            using (var ms = new MemoryStream())
            {
                var pdfExport = new PDFExport
                {
                    Name = "Exported",
                    Background = true
                };

                report1.Export(pdfExport, ms);

                //設定本機資料夾 
                string DirectoryNAME = SETPATHFLODER();
                File.WriteAllBytes(DirectoryNAME+"Exported.pdf", ms.ToArray());
            }


        }

        public string SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            
                    
            FASTSQL.AppendFormat(@"  
                                SELECT *
                                ,CASE WHEN TC018='1' THEN '應稅內含' WHEN TC018='2' THEN '應稅外加' WHEN TC018='3' THEN '零稅率' WHEN TC018='4' THEN '免稅 'WHEN TC018='9' THEN '不計稅' END AS TC018NAME
                                FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.CMSMQ,[TK].dbo.PURMA,[TK].dbo.CMSMV
                                WHERE TC001=TD001 AND TC002=TD002
                                AND MQ001=TC001
                                AND TC004=MA001
                                AND TC011=MV001
                                AND TD002='20220623003' 
                                ");

            return FASTSQL.ToString();
        }

        //設定本機資料夾 
        public string  SETPATHFLODER()
        {
            string DirectoryNAME = null;
            string DATES = DateTime.Now.ToString("yyyyMMdd");

            DirectoryNAME = @"C:\PDFTEMP\" + DATES.ToString() + @"\";

            //不存在就新增資料夾
            if (Directory.Exists(DirectoryNAME))
            {                

            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }

            return DirectoryNAME;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        #endregion
    }
}
