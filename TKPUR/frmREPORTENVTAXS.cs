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
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKPUR
{
    public partial class frmREPORTENVTAXS : Form
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
        DataSet ds2 = new DataSet();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string SALSESID = null;
        int result;

        public Report report1 { get; private set; }

        public frmREPORTENVTAXS()
        {
            InitializeComponent();
        }

        #region FUNCTION
       
        private void frmREPORTENVTAXS_Load(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            int month = now.Month;
            int year = now.Year;

            // 计算当前月份属于哪个双月区间
            int startMonth = ((month - 1) / 2) * 2 + 1; // 1, 3, 5, 7, 9, 11
            int endMonth = startMonth + 1;              // 2, 4, 6, 8, 10, 12

            DateTime lastMonth = new DateTime(year, startMonth, 1);
            DateTime nowMonth = new DateTime(year, endMonth, 1);

            dateTimePicker2.Value = lastMonth;
            dateTimePicker3.Value = nowMonth;

        }
        public void SETDATE()
        {

        }
        public void SETFASTREPORT(string YYYYMM)
        {
            string YY = YYYYMM.Substring(0,4);
            string MM = YYYYMM.Substring(4,2);

            string SQL;
            string SQL1;
            report1 = new Report();
            report1.Load(@"REPORT\環保稅.frx");

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
            SQL = SETFASETSQL(YY,MM);
            Table.SelectCommand = SQL;
            TableDataSource Table1 = report1.GetDataSource("Table1") as TableDataSource;
            SQL1= SETFASETSQL1(YY, MM);
            Table1.SelectCommand = SQL1;
            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string YY,string MM)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            StringBuilder STRQUERY2 = new StringBuilder();

            DataTable DT = FIND_TKCOPMATAXSMB001PUR();
            if(DT!=null&& DT.Rows.Count>=1)
            {
                STRQUERY.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY.AppendFormat(@" TH004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY.AppendFormat(@" OR");
                    }
                }

                STRQUERY.AppendFormat(@" )");
            }

            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY2.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY2.AppendFormat(@" TI004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY2.AppendFormat(@" OR");
                    }
                }

                STRQUERY2.AppendFormat(@" )");
            }


            FASTSQL.AppendFormat(@"  
                                SELECT *
                                FROM 
                                (
                                SELECT SUBSTRING(TG003,1,4) AS '年',SUBSTRING(TG003,5,2)  AS '月',TG005 AS '廠商代',MA002 AS '廠商',MA005 AS '統編',TH004 AS '品號',MB002 AS '品名',SUM(TH015)  AS '進貨驗收數量',TH008 AS '單位'
                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH, [TK].dbo.PURMA, [TK].dbo.INVMB 
                                WHERE TG001=TH001 AND TG002=TH002
                                AND MA001=TG005
                                AND MB001=TH004
                                AND TG013='Y'
                                AND  {2} 
                                AND SUBSTRING(TG003,1,4)='{0}'
                                AND SUBSTRING(TG003,5,2)='{1}'
                                GROUP BY  SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG005,MA002,MA005,TH004,MB002,MB003,TH008
                               
                                
                                UNION ALL
                                SELECT SUBSTRING(TH003,1,4) AS '年',SUBSTRING(TH003,5,2)  AS '月',TH005 AS '廠商代',MA002 AS '廠商',MA005 AS '統編',TI004 AS '品號',MB002 AS '品名',SUM(TI007)  AS '進貨驗收數量',TI008 AS '單位'
                                FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI, [TK].dbo.PURMA, [TK].dbo.INVMB 
                                WHERE TH001=TI001 AND TH002=TI002
                                AND MA001=TH005
                                AND MB001=TI004
                                AND TH023='Y'
                                AND  {3} 
                                AND SUBSTRING(TH003,1,4)='{0}'
                                AND SUBSTRING(TH003,5,2)='{1}'
                                GROUP BY  SUBSTRING(TH003,1,4),SUBSTRING(TH003,5,2),TH005,MA002,MA005,TI004,MB002,MB003,TI008
                                
                                )  AS TEMP
                                WHERE 1=1
                                ORDER BY 年,月,廠商代,廠商,統編,品號,品名,單位
                                    ", YY,MM, STRQUERY.ToString(), STRQUERY2.ToString());

            return FASTSQL.ToString();
        }

        public string SETFASETSQL1(string YY, string MM)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            DataTable DT = FIND_TKCOPMATAXSMB001COP();
            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY.AppendFormat(@" MD003 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY.AppendFormat(@" OR");
                    }
                }

                STRQUERY.AppendFormat(@" )");
            }

            FASTSQL.AppendFormat(@"                                
                                SELECT  SUBSTRING(TG003,1,4) AS '年',SUBSTRING(TG003,5,2)  AS '月',TG004 AS '客戶代',MA002 AS '客戶',MA010 AS '統編',TH004 ,MB1.MB002 ,SUM(LA011),MB1.MB004,MC004,MD006,MD007,MD003 AS '品號',MB2.MB002 AS '品名',SUM(CONVERT(DECIMAL(16,0),(LA011/MD006*MD007*MC004)))  AS '數量',MB2.MB004 AS '單位'
                                FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB MB1,[TK].dbo.COPMA,[TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB MB2
                                WHERE TG001=TH001 AND TG002=TH002
                                AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                                AND TH004=MB1.MB001
                                AND TG004=MA001
                                AND MC001=TH004
                                AND MC001=MD001
                                AND MD003=MB2.MB001
                                AND {2}
                                AND MD035 NOT LIKE '%蓋%'
                                AND (TG004 LIKE '2%' OR TG004 LIKE 'A%')
                                AND TG004 IN (SELECT  [MA001] FROM [TKPUR].[dbo].[TKCOPMATAXS])
                                AND SUBSTRING(TG003,1,4)='{0}' 
                                AND SUBSTRING(TG003,5,2)='{1}'
                                GROUP BY SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG004,MA002,MA010,TH004,MB1.MB002,MB1.MB004,MC004,MD006,MD007,MD003,MB2.MB002,MB2.MB004
                                ORDER BY SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG004,MA002,MA010,TH004

                                    ", YY, MM, STRQUERY.ToString());

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2(string STARTYYYYMM,string ENDYYYYMM)
        {
            string STARTYY = STARTYYYYMM.Substring(0, 4);
            string STARTMM = STARTYYYYMM.Substring(4, 2);
            string ENDYY = ENDYYYYMM.Substring(0, 4);
            string ENDMM = ENDYYYYMM.Substring(4, 2);

            string SQL;
            string SQL1;
            report1 = new Report();
            report1.Load(@"REPORT\環保稅.frx");

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
            SQL = SETFASETSQL2A(STARTYY, STARTMM, ENDYY, ENDMM);
            Table.SelectCommand = SQL;
            TableDataSource Table1 = report1.GetDataSource("Table1") as TableDataSource;
            SQL1 = SETFASETSQL2B(STARTYY, STARTMM, ENDYY, ENDMM);
            Table1.SelectCommand = SQL1;
            report1.Preview = previewControl1;
            report1.Show();

        }
        public string SETFASETSQL2A(string STARTYY, string STARTMM, string ENDYY, string ENDMM)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();
            StringBuilder STRQUERY2 = new StringBuilder();

            DataTable DT = FIND_TKCOPMATAXSMB001PUR();
            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY.AppendFormat(@" TH004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY.AppendFormat(@" OR");
                    }
                }

                STRQUERY.AppendFormat(@" )");
            }

            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY2.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY2.AppendFormat(@" TI004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY2.AppendFormat(@" OR");
                    }
                }

                STRQUERY2.AppendFormat(@" )");
            }

            FASTSQL.AppendFormat(@"  
                                
                                SELECT 年, '{0}月' 月,廠商代,廠商,統編,品號,品名,SUM(進貨驗收數量) 進貨驗收數量,單位
                                FROM
                                (
                                SELECT SUBSTRING(TG003,1,4) AS '年',SUBSTRING(TG003,5,2)  AS '月',TG005 AS '廠商代',MA002 AS '廠商',MA005 AS '統編',TH004 AS '品號',MB002 AS '品名',SUM(TH015)  AS '進貨驗收數量',TH008 AS '單位'
                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH, [TK].dbo.PURMA, [TK].dbo.INVMB 
                                WHERE TG001=TH001 AND TG002=TH002
                                AND MA001=TG005
                                AND MB001=TH004
                                AND TG013='Y'
                                AND  {4}
                                AND SUBSTRING(TG003,1,4)='{1}'
                                AND SUBSTRING(TG003,5,2)>='{2}'
                                AND SUBSTRING(TG003,5,2)<='{3}'
                                GROUP BY  SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG005,MA002,MA005,TH004,MB002,MB003,TH008
                                
                                UNION ALL
                                SELECT SUBSTRING(TH003,1,4) AS '年',SUBSTRING(TH003,5,2)  AS '月',TH005 AS '廠商代',MA002 AS '廠商',MA005 AS '統編',TI004 AS '品號',MB002 AS '品名',SUM(TI007)  AS '進貨驗收數量',TI008 AS '單位'
                                FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI, [TK].dbo.PURMA, [TK].dbo.INVMB 
                                WHERE TH001=TI001 AND TH002=TI002
                                AND MA001=TH005
                                AND MB001=TI004
                                AND TH023='Y'
                                AND  {5} 
                                AND SUBSTRING(TH003,1,4)='{1}'
                                AND SUBSTRING(TH003,5,2)>='{2}'
                                AND SUBSTRING(TH003,5,2)<='{3}'
                                GROUP BY  SUBSTRING(TH003,1,4),SUBSTRING(TH003,5,2),TH005,MA002,MA005,TI004,MB002,MB003,TI008
                                
                                ) AS TEMP
                                GROUP BY 年, 廠商代,廠商,統編,品號,品名,單位
                                ORDER BY 年, 廠商代,廠商,統編,品號,品名,單位
     
                                    ", STARTMM+"-"+ ENDMM, STARTYY, STARTMM, ENDMM, STRQUERY.ToString(), STRQUERY2.ToString());

            return FASTSQL.ToString();
        }
        public string SETFASETSQL2B(string STARTYY, string STARTMM, string ENDYY, string ENDMM)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            DataTable DT = FIND_TKCOPMATAXSMB001COP();
            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY.AppendFormat(@" MD003 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY.AppendFormat(@" OR");
                    }
                }

                STRQUERY.AppendFormat(@" )");
            }

            FASTSQL.AppendFormat(@"                                
                              
                                    SELECT 年, '{0}月' 月,客戶代,客戶,統編,品號,品名,SUM(數量) 數量,單位
                                    FROM
                                    (
                                    SELECT  SUBSTRING(TG003,1,4) AS '年',SUBSTRING(TG003,5,2)  AS '月',TG004 AS '客戶代',MA002 AS '客戶',MA010 AS '統編',TH004 ,MB1.MB002 ,SUM(LA011) LA011,MB1.MB004,MC004,MD006,MD007,MD003 AS '品號',MB2.MB002 AS '品名',SUM(CONVERT(DECIMAL(16,0),(LA011/MD006*MD007*MC004)))  AS '數量',MB2.MB004 AS '單位'
                                    FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB MB1,[TK].dbo.COPMA,[TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB MB2
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                                    AND TH004=MB1.MB001
                                    AND TG004=MA001
                                    AND MC001=TH004
                                    AND MC001=MD001
                                    AND MD003=MB2.MB001
                                    AND {4}
                                    AND MD035 NOT LIKE '%蓋%'
                                    AND (TG004 LIKE '2%' OR TG004 LIKE 'A%')
                                    AND TG004 IN (SELECT  [MA001] FROM [TKPUR].[dbo].[TKCOPMATAXS])
                                    AND SUBSTRING(TG003,1,4)='{1}' 
                                    AND SUBSTRING(TG003,5,2)>='{2}'
                                    AND SUBSTRING(TG003,5,2)>='{3}'
                                    GROUP BY SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG004,MA002,MA010,TH004,MB1.MB002,MB1.MB004,MC004,MD006,MD007,MD003,MB2.MB002,MB2.MB004
                                    ) AS TEMP
                                    GROUP BY 年, 客戶代,客戶,統編,品號,品名,單位
                                    ORDER BY 年, 客戶代,客戶,統編,品號,品名,單位

                                    ", STARTMM + "-" + ENDMM, STARTYY, STARTMM, ENDMM, STRQUERY.ToString());

            return FASTSQL.ToString();
        }

        public DataTable FIND_TKCOPMATAXSMB001PUR()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

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
                                   SELECT  [MB001]
                                    FROM [TKPUR].[dbo].[TKCOPMATAXSMB001PUR]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        public DataTable FIND_TKCOPMATAXSMB001COP()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

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
                                   SELECT  [MB001]
                                    FROM [TKPUR].[dbo].[TKCOPMATAXSMB001COP]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                if (ds.Tables["ds"].Rows.Count >= 1)
                {
                    return ds.Tables["ds"];
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void Search_TKTAXCODES()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

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
                                    [CODES] AS '材質細碼'
                                    ,[VOLUMES] AS '容積'
                                    ,[WEIGHTS] AS '容器本體'
                                    ,[OTHERWEIGHTS] AS '附件'
                                    ,[RATES] AS '費率'
                                    FROM [TKPUR].[dbo].[TKTAXCODES]
                                    ORDER BY [CODES],[VOLUMES]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView1.DataSource = ds.Tables["TEMPds1"];
                    dataGridView1.AutoResizeColumns();

                }
                else
                {
                    dataGridView1.DataSource = null;
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SET_NULL();

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox6.Text = row.Cells["材質細碼"].Value.ToString();
                    textBox7.Text = row.Cells["容積"].Value.ToString();
                    textBox8.Text = row.Cells["容器本體"].Value.ToString();
                    textBox9.Text = row.Cells["附件"].Value.ToString();
                    textBox10.Text = row.Cells["費率"].Value.ToString();

                }
                else
                {
                   
                }
            }
        }

        public void Search_TKTAXCODES_V2()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

            

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
                                    [CODES] AS '材質細碼'
                                    ,[VOLUMES] AS '容積'
                                    ,[WEIGHTS] AS '容器本體'
                                    ,[OTHERWEIGHTS] AS '附件'
                                    ,[RATES] AS '費率'
                                    FROM [TKPUR].[dbo].[TKTAXCODES]
                                    ORDER BY [CODES],[VOLUMES]
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds.Tables["TEMPds1"];
                    dataGridView2.AutoResizeColumns();

                }
                else
                {
                    dataGridView2.DataSource = null;
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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            SET_NULL_V2();

            string CODES = "";
            string VOLUMES = "";
            string WEIGHTS = "";
            string OTHERWEIGHTS = "";
            string RATES = "";

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;


                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    CODES = row.Cells["材質細碼"].Value.ToString();
                    VOLUMES = row.Cells["容積"].Value.ToString();
                    WEIGHTS = row.Cells["容器本體"].Value.ToString();
                    OTHERWEIGHTS = row.Cells["附件"].Value.ToString();
                    RATES = row.Cells["費率"].Value.ToString();

                    textBox13.Text = CODES;
                    textBox14.Text = VOLUMES;
                    textBox15.Text = WEIGHTS;
                    textBox16.Text = OTHERWEIGHTS;
                    textBox17.Text = RATES;

                    Search_TKTAXCODESMB001(CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS);
                }
                else
                {

                }
            }
        }

        public void Search_TKTAXCODESMB001(string CODES, string VOLUMES, string WEIGHTS, string OTHERWEIGHTS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

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
                                    [MB001] AS '品號'
                                    ,[MB002] AS '品名'
                                    FROM [TKPUR].[dbo].[TKTAXCODESMB001]
                                    WHERE 
                                    [CODES]='{0}'
                                    AND [VOLUMES]={1}
                                    AND [WEIGHTS]={2}
                                    AND [OTHERWEIGHTS]={3}
                                    ORDER BY [CODES],[MB001]
                                    ", CODES,  VOLUMES,  WEIGHTS,  OTHERWEIGHTS);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds.Tables["TEMPds1"];
                    dataGridView3.AutoResizeColumns();

                }
                else
                {
                    dataGridView3.DataSource = null;
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            SET_NULL_V3();

            string CODES = "";
            string VOLUMES = "";
            string WEIGHTS = "";
            string OTHERWEIGHTS = "";
            string RATES = "";
            string MB001 = "";
            string MB002 = "";

            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    MB001 = row.Cells["品號"].Value.ToString();
                    MB002 = row.Cells["品名"].Value.ToString();

                    CODES = textBox13.Text.Trim();
                    VOLUMES = textBox14.Text.Trim();
                    WEIGHTS = textBox15.Text.Trim();
                    OTHERWEIGHTS = textBox16.Text.Trim();
                    RATES = textBox17.Text.Trim();

                    textBox18.Text = MB001;
                    textBox19.Text = MB002;
                }
                else
                {

                }
            }
        }

        public void ADD_TKTAXCODES(string CODES, string VOLUMES, string WEIGHTS, string OTHERWEIGHTS, string RATES)
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

               
                sbSql.AppendFormat(@" 
                                    INSERT INTO  [TKPUR].[dbo].[TKTAXCODES]
                                    (
                                    [CODES]
                                    ,[VOLUMES]
                                    ,[WEIGHTS]
                                    ,[OTHERWEIGHTS]
                                    ,[RATES]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,{1}
                                    ,{2}
                                    ,{3}
                                    ,{4}
                                    )
                                    ", CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES);


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
            catch(Exception  ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELETE_TKTAXCODES(string CODES, string VOLUMES, string WEIGHTS, string OTHERWEIGHTS, string RATES)
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@" 
                                    DELETE [TKPUR].[dbo].[TKTAXCODES]
                                    WHERE  [CODES]='{0}'AND [VOLUMES]={1} AND [WEIGHTS]={2} AND [OTHERWEIGHTS]={3} AND [RATES]={4}
                                   
                                    ", CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES);


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            string MB001= textBox11.Text.Trim();
            if (!string.IsNullOrEmpty(MB001))
            {
                DataTable DT = SEARCH_INVMB(MB001);

                if(DT!=null && DT.Rows.Count>=1)
                {
                    textBox12.Text = DT.Rows[0]["MB002"].ToString().Trim();
                }
            }
           
        }

        public DataTable SEARCH_INVMB(string MB001)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
            DataSet ds = new DataSet();

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
                                    MB001,MB002
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001='{0}'
                                    ", MB001);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return (ds.Tables["TEMPds1"]);
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADD_TKTAXCODESMB001(
            string CODES ,
            string VOLUMES,
            string WEIGHTS ,
            string OTHERWEIGHTS ,
            string RATES ,
            string MB001,
            string MB002
            )
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"                                   
                                    INSERT INTO [TKPUR].[dbo].[TKTAXCODESMB001]
                                    (
                                    [CODES]
                                    ,[VOLUMES]
                                    ,[WEIGHTS]
                                    ,[OTHERWEIGHTS]
                                    ,[RATES]
                                    ,[MB001]
                                    ,[MB002]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,{1}
                                    ,{2}
                                    ,{3}
                                    ,{4}
                                    ,'{5}'
                                    ,'{6}'


                                    )

                                   
                                    ", CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES, MB001, MB002);


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELETE_TKTAXCODESMB001(
                  string CODES,
                  string VOLUMES,
                  string WEIGHTS,
                  string OTHERWEIGHTS,
                  string RATES,
                  string MB001,
                  string MB002
              )
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"                                   
                                   DELETE [TKPUR].[dbo].[TKTAXCODESMB001]
                                   WHERE 
                                    [CODES]='{0}'
                                    AND [VOLUMES]={1}
                                    AND [WEIGHTS]={2}
                                    AND [OTHERWEIGHTS]={3}
                                    AND [RATES]={4}
                                    AND [MB001]='{5}'
                                    AND [MB002]='{6}'
                                   

                                   
                                    ", CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES, MB001, MB002);


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void ADD_TKTAXREPORTCOP_TKTAXREPORTPUR(string STARTYYYYMM, string ENDYYYYMM)
        {
            string STARTYY = STARTYYYYMM.Substring(0, 4);
            string STARTMM = STARTYYYYMM.Substring(4, 2);
            string ENDYY = ENDYYYYMM.Substring(0, 4);
            string ENDMM = ENDYYYYMM.Substring(4, 2);

            StringBuilder STRQUERY1 = new StringBuilder();
            StringBuilder STRQUERY2 = new StringBuilder();
            StringBuilder STRQUERY3 = new StringBuilder();

            DataTable DT = FIND_TKCOPMATAXSMB001PUR();
            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY1.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY1.AppendFormat(@" TH004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY1.AppendFormat(@" OR");
                    }
                }

                STRQUERY1.AppendFormat(@" )");
            }

            if (DT != null && DT.Rows.Count >= 1)
            {
                STRQUERY3.AppendFormat(@" (");
                int rowCount = DT.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY3.AppendFormat(@" TI004 LIKE '{0}%'", DT.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY3.AppendFormat(@" OR");
                    }
                }

                STRQUERY3.AppendFormat(@" )");
            }

            DataTable DT2 = FIND_TKCOPMATAXSMB001COP();
            if (DT2 != null && DT2.Rows.Count >= 1)
            {
                STRQUERY2.AppendFormat(@" (");
                int rowCount = DT2.Rows.Count;

                for (int i = 0; i < rowCount; i++)
                {
                    STRQUERY2.AppendFormat(@" MD003 LIKE '{0}%'", DT2.Rows[i]["MB001"].ToString());

                    // 在最後一個元素之後不添加 "OR"
                    if (i < rowCount - 1)
                    {
                        STRQUERY2.AppendFormat(@" OR");
                    }
                }

                STRQUERY2.AppendFormat(@" )");
            }
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    DELETE [TKPUR].[dbo].[TKTAXREPORTPUR]
                                    INSERT INTO [TKPUR].[dbo].[TKTAXREPORTPUR]
                                    (
                                    [年]
                                    ,[月]
                                    ,[廠商代]
                                    ,[廠商]
                                    ,[統編]

                                    ,[品號]
                                    ,[品名]
                                    ,[進貨驗收數量]
                                    ,[單位]
                                    ,[材質細碼]
                                    ,[容積]
                                    ,[容器本體]
                                    ,[附件]
                                    ,[費率]
                                    ,[MB001]
                                    ,[MB002]
                                    )
                                    SELECT 
                                    年,月,廠商代,廠商,統編,品號,品名,進貨驗收數量,單位
                                    ,[CODES] AS '材質細碼'
                                    ,[VOLUMES] AS '容積'
                                    ,[WEIGHTS] AS '容器本體'
                                    ,[OTHERWEIGHTS] AS '附件'
                                    ,[RATES] AS '費率'
                                    ,[MB001]
                                    ,[MB002]
                                    FROM 
                                    (
                                    SELECT SUBSTRING(TG003,1,4) AS '年',SUBSTRING(TG003,5,2)  AS '月',TG005 AS '廠商代',MA002 AS '廠商',MA005 AS '統編',TH004 AS '品號',MB002 AS '品名',SUM(TH015)  AS '進貨驗收數量',TH008 AS '單位'
                                    FROM [TK].dbo.PURTG,[TK].dbo.PURTH, [TK].dbo.PURMA, [TK].dbo.INVMB 
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND MA001=TG005
                                    AND MB001=TH004
                                    AND TG013='Y'
                                    AND {3}
                                    AND SUBSTRING(TG003,1,4)='{0}'
                                    AND SUBSTRING(TG003,5,2)>='{1}'
                                    AND SUBSTRING(TG003,5,2)<='{2}'
                                    GROUP BY  SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG005,MA002,MA005,TH004,MB002,MB003,TH008

                                    UNION ALL
                                    SELECT SUBSTRING(TH003,1,4) AS '年',SUBSTRING(TH003,5,2)  AS '月',TH005 AS '廠商代',MA002 AS '廠商',MA005 AS '統編',TI004 AS '品號',MB002 AS '品名',SUM(TI007)  AS '進貨驗收數量',TI008 AS '單位'
                                    FROM [TK].dbo.MOCTH,[TK].dbo.MOCTI, [TK].dbo.PURMA, [TK].dbo.INVMB 
                                    WHERE TH001=TI001 AND TH002=TI002
                                    AND MA001=TH005
                                    AND MB001=TI004
                                    AND TH023='Y'
                                    AND  {4} 
                                    AND SUBSTRING(TH003,1,4)='{0}'
                                    AND SUBSTRING(TH003,5,2)>='{1}'
                                    AND SUBSTRING(TH003,5,2)<='{2}'
                                    GROUP BY  SUBSTRING(TH003,1,4),SUBSTRING(TH003,5,2),TH005,MA002,MA005,TI004,MB002,MB003,TI008
                                
                                    ) AS TEMP 
                                    LEFT JOIN [TKPUR].[dbo].[TKTAXCODESMB001] ON [TKTAXCODESMB001].MB001=TEMP.品號
                                    ORDER BY  年,月,廠商代,廠商,統編,品號,品名,單位
                                   
                                    ", STARTYY, STARTMM, ENDMM, STRQUERY1.ToString(), STRQUERY3.ToString());

                sbSql.AppendFormat(@"  ");

                sbSql.AppendFormat(@"  
                                    DELETE  [TKPUR].[dbo].[TKTAXREPORTCOP]
                                    INSERT INTO [TKPUR].[dbo].[TKTAXREPORTCOP]
                                    (
                                     [年]
                                    ,[月]
                                    ,[客戶代]
                                    ,[客戶]
                                    ,[統編]
                                    ,[主品號]
                                    ,[主品名]
                                    ,[主銷售數量]
                                    ,[品號]
                                    ,[品名]
                                    ,[進貨驗收數量]
                                    ,[單位]
                                    ,[材質細碼]
                                    ,[容積]
                                    ,[容器本體]
                                    ,[附件]
                                    ,[費率]
                                    ,[MB001]
                                    ,[MB002]
                                    )
                                    SELECT 
                                     [年]
                                    ,[月]
                                    ,[客戶代]
                                    ,[客戶]
                                    ,[統編]
                                    ,[主品號]
                                    ,[主品名]
                                    ,[主銷售數量]
                                    ,[品號]
                                    ,[品名]
                                    ,[數量]
                                    ,[單位]
                                    ,[CODES] AS '材質細碼'
                                    ,[VOLUMES] AS '容積'
                                    ,[WEIGHTS] AS '容器本體'
                                    ,[OTHERWEIGHTS] AS '附件'
                                    ,[RATES] AS '費率'
                                    ,[MB001]
                                    ,[MB002]
                                    FROM 
                                    (
                                    SELECT  SUBSTRING(TG003,1,4) AS '年',SUBSTRING(TG003,5,2)  AS '月',TG004 AS '客戶代',MA002 AS '客戶',MA010 AS '統編',TH004  AS '主品號',MB1.MB002 AS '主品名' ,SUM(LA011)  AS '主銷售數量',MB1.MB004,MC004,MD006,MD007,MD003 AS '品號',MB2.MB002 AS '品名',SUM(CONVERT(DECIMAL(16,0),(LA011/MD006*MD007*MC004)))  AS '數量',MB2.MB004 AS '單位'
                                    FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB MB1,[TK].dbo.COPMA,[TK].dbo.BOMMC,[TK].dbo.BOMMD,[TK].dbo.INVMB MB2
                                    WHERE TG001=TH001 AND TG002=TH002
                                    AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                                    AND TH004=MB1.MB001
                                    AND TG004=MA001
                                    AND MC001=TH004
                                    AND MC001=MD001
                                    AND MD003=MB2.MB001
                                    AND {3}
                                    AND MD035 NOT LIKE '%蓋%'
                                    AND (TG004 LIKE '2%' OR TG004 LIKE 'A%')
                                    AND TG004 IN (SELECT  [MA001] FROM [TKPUR].[dbo].[TKCOPMATAXS])
                                    AND SUBSTRING(TG003,1,4)='{0}' 
                                    AND SUBSTRING(TG003,5,2)>='{1}'
                                    AND SUBSTRING(TG003,5,2)<='{2}'
                                    GROUP BY SUBSTRING(TG003,1,4),SUBSTRING(TG003,5,2),TG004,MA002,MA010,TH004,MB1.MB002,MB1.MB004,MC004,MD006,MD007,MD003,MB2.MB002,MB2.MB004
                                    )  AS TEMP
                                    LEFT JOIN [TKPUR].[dbo].[TKTAXCODESMB001] ON [TKTAXCODESMB001].MB001=TEMP.品號
                                    ORDER BY  年,月,客戶代,客戶,統編,品號,品名,單位

                                    ", STARTYY, STARTMM, ENDMM, STRQUERY2.ToString());


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void DELETE_TKTAXCODESMB001_CODE(
                  string CODES,
                  string VOLUMES,
                  string WEIGHTS,
                  string OTHERWEIGHTS,
                  string RATES
            )
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

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"                                   
                                   DELETE [TKPUR].[dbo].[TKTAXCODESMB001]
                                   WHERE 
                                    [CODES]='{0}'
                                    AND [VOLUMES]={1}
                                    AND [WEIGHTS]={2}
                                    AND [OTHERWEIGHTS]={3}
                                    AND [RATES]={4}
                          
                                   

                                   
                                    ", CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES);


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
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SET_NULL()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";

        }

        public void SET_NULL_V2()
        {
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
        }

        public void SET_NULL_V3()
        {
            textBox18.Text = "";
            textBox19.Text = "";
      
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMM"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker2.Value.ToString("yyyyMM"), dateTimePicker3.Value.ToString("yyyyMM"));
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Search_TKTAXCODES();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string CODES = textBox1.Text.Trim();
            string VOLUMES = textBox2.Text.Trim();
            string WEIGHTS = textBox3.Text.Trim();
            string OTHERWEIGHTS = textBox4.Text.Trim();
            string RATES = textBox5.Text.Trim();

            if(!string.IsNullOrEmpty(CODES))
            {
                ADD_TKTAXCODES(CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES);

                SET_NULL();
                Search_TKTAXCODES();
            }

          
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string CODES = textBox6.Text.Trim();
            string VOLUMES = textBox7.Text.Trim();
            string WEIGHTS = textBox8.Text.Trim();
            string OTHERWEIGHTS = textBox9.Text.Trim();
            string RATES = textBox10.Text.Trim();

            DialogResult dialogResult = MessageBox.Show("要刪除了?" + Environment.NewLine + "注意:如果刪除材質細碼，會連同品號關係一並刪除", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(CODES))
                {
                    DELETE_TKTAXCODES(CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS, RATES);

                    DELETE_TKTAXCODESMB001_CODE(
                      CODES,
                      VOLUMES,
                      WEIGHTS,
                      OTHERWEIGHTS,
                      RATES                    
                    );

                    SET_NULL();
                    Search_TKTAXCODES();
                }

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Search_TKTAXCODES_V2();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string CODES = textBox13.Text.Trim();
            string VOLUMES = textBox14.Text.Trim();
            string WEIGHTS = textBox15.Text.Trim();
            string OTHERWEIGHTS = textBox16.Text.Trim();
            string RATES = textBox17.Text.Trim();
            string MB001 = textBox11.Text.Trim();
            string MB002 = textBox12.Text.Trim();

            if(!string.IsNullOrEmpty(MB001) && !string.IsNullOrEmpty(MB002))
            {
                ADD_TKTAXCODESMB001(
                  CODES,
                  VOLUMES,
                  WEIGHTS,
                  OTHERWEIGHTS,
                  RATES,
                  MB001,
                  MB002
              );

                Search_TKTAXCODESMB001(CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS);
            }
          
        }




     

        private void button8_Click(object sender, EventArgs e)
        {
            string CODES = textBox13.Text.Trim();
            string VOLUMES = textBox14.Text.Trim();
            string WEIGHTS = textBox15.Text.Trim();
            string OTHERWEIGHTS = textBox16.Text.Trim();
            string RATES = textBox17.Text.Trim();

            string MB001 = textBox18.Text.Trim();
            string MB002 = textBox19.Text.Trim();
           

            DialogResult dialogResult = MessageBox.Show("要刪除了?" , "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(MB001) && !string.IsNullOrEmpty(MB002))
                {
                    DELETE_TKTAXCODESMB001(
                      CODES,
                      VOLUMES,
                      WEIGHTS,
                      OTHERWEIGHTS,
                      RATES,
                      MB001,
                      MB002
                    );
                    Search_TKTAXCODESMB001(CODES, VOLUMES, WEIGHTS, OTHERWEIGHTS);
                }

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            ADD_TKTAXREPORTCOP_TKTAXREPORTPUR(dateTimePicker4.Value.ToString("yyyyMM"), dateTimePicker5.Value.ToString("yyyyMM"));
        }
        #endregion
    }
}
