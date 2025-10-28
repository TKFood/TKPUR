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

        private void frmREPORTPAYS_Load(object sender, EventArgs e)
        {
            AddCheckBoxColumn();
            AddCheckBoxColumn2();
        }

        private void AddCheckBoxColumn()
        {
            // 1. 建立 DataGridViewCheckBoxColumn 實例
            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();

            // 2. 設定欄位屬性
            checkBoxColumn.HeaderText = "選取"; // 欄位標題
            checkBoxColumn.Name = "SelectedCheckbox"; // 欄位名稱 (建議設定，方便後續程式碼存取)
            checkBoxColumn.ValueType = typeof(bool); // 設定儲存的值的型別為 bool
                                                     // 可選：設定寬度或自動調整模式
            checkBoxColumn.Width = 50;
            checkBoxColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            // 3. 將欄位新增到 DataGridView 的 Columns 集合中
            // 預設是加在最後一欄，如果您想放在第一欄，可以使用 Insert(0, checkBoxColumn)
            this.dataGridView1.Columns.Insert(0, checkBoxColumn);
            // 或者 this.dataGridView1.Columns.Add(checkBoxColumn); // 加到最後
        }
        private void AddCheckBoxColumn2()
        {
            // 1. 建立 DataGridViewCheckBoxColumn 實例
            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();

            // 2. 設定欄位屬性
            checkBoxColumn.HeaderText = "選取"; // 欄位標題
            checkBoxColumn.Name = "SelectedCheckbox"; // 欄位名稱 (建議設定，方便後續程式碼存取)
            checkBoxColumn.ValueType = typeof(bool); // 設定儲存的值的型別為 bool
                                                     // 可選：設定寬度或自動調整模式
            checkBoxColumn.Width = 50;
            checkBoxColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            // 3. 將欄位新增到 DataGridView 的 Columns 集合中
            // 預設是加在最後一欄，如果您想放在第一欄，可以使用 Insert(0, checkBoxColumn)
            this.dataGridView2.Columns.Insert(0, checkBoxColumn);
            // 或者 this.dataGridView1.Columns.Add(checkBoxColumn); // 加到最後
        }
        #region FUNCTION

        public void Search( string TG002, string MA001)
        {
            DataSet ds = new DataSet();

            StringBuilder sbSqlQuery1 = new StringBuilder();
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

                if (!string.IsNullOrEmpty(TG002))
                {
                    sbSqlQuery1.AppendFormat(@" 
                                            AND 單號 LIKE '%{0}%'
                                                ", TG002);
                }
                else
                {
                    sbSqlQuery1.AppendFormat(@" 
                                           
                                                ");
                }


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

                             

                //採購的進貨+製令的託外進貨
                sbSql.AppendFormat(@"                                    
                                   SELECT *
                                    FROM 
                                    (
                                    SELECT 
                                    TG001 AS '單別'
                                    ,TG002 AS '單號'
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

                                    ,TG005 AS '供應廠商'
                                    ,TG033
                                    FROM [TK].dbo.PURTG
                                    WHERE 1=1
                                    ) AS TEMP
                                    WHERE 1=1   
                                    AND TG033 IN (
	                                    SELECT  [TG033]
	                                    FROM [TKPUR].[dbo].[TKPURTGTG033]
                                    )                                
                                    {0}
                                    {1}
                                    ORDER BY 單別,單號
                               
                                    ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString());

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

        /// <summary>
        /// 找出 DataGridView 中被勾選的行的名稱。
        /// </summary>
        /// <param name="dataGridView">要操作的 DataGridView 控制項。</param>
        /// <param name="checkBoxColumnName">CheckBox 欄位的 Name 屬性值 (例如: "SelectedCheckbox")。</param>
        /// <param name="nameColumnName">包含名稱資料的欄位 Name 屬性值 (例如: "Name")。</param>
        /// <returns>一個包含所有被勾選名稱的 List<string>。</returns>
        public List<string> GetSelectedNamesFromDGV(DataGridView dataGridView, string checkBoxColumnName)
        {
            List<string> selectedNames = new List<string>();

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // 忽略新增行 (New Row)
                if (row.IsNewRow)
                {
                    continue;
                }

                // 確保兩個目標欄位都存在
                if (row.Cells[checkBoxColumnName] != null && row.Cells["單號"] != null)
                {
                    // 獲取 CheckBox 欄位的值
                    object checkBoxValue = row.Cells[checkBoxColumnName].Value;
                    bool isChecked = false;

                    // 嘗試將值轉換為 bool。需要處理 null 或 DBNull 的情況。
                    if (checkBoxValue != null && checkBoxValue != DBNull.Value)
                    {
                        try
                        {
                            // 2. 使用 Convert.ToBoolean 進行轉換
                            // 它能處理 bool, 1/0 (int), "True"/"False" (string) 等情況
                            isChecked = Convert.ToBoolean(checkBoxValue);
                        }
                        catch (InvalidCastException)
                        {
                            // 處理轉換失敗的情況 (例如欄位包含非 1/0 的文字)
                            // 您可以記錄錯誤或設定預設值
                            isChecked = false;
                        }
                        catch (FormatException)
                        {
                            // 處理格式錯誤 (例如欄位是無效字串)
                            isChecked = false;
                        }
                    }

                    // 如果 CheckBox 被勾選，則獲取對應的名稱
                    if (isChecked)
                    {
                        // 獲取名稱欄位的值 (通常是 string)
                        object TG001 = row.Cells["單別"].Value;
                        object TG002 = row.Cells["單號"].Value;
                        if (TG002 != null && TG002 != DBNull.Value)
                        {
                            selectedNames.Add(TG001.ToString().Trim()+TG002.ToString().Trim());
                        }
                    }
                }
            }

            return selectedNames;
        }

        public void SETFASTREPORT(string sqlInCondition)
        {
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();
            report1.Load(@"REPORT\請款憑單BY進貨單.frx");

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

            SQL = SETFASETSQL1(sqlInCondition);

            Table.SelectCommand = SQL.ToString(); ;

            //report1.SetParameterValue("P1", COMMENT);

            report1.Preview = previewControl1;
            report1.Show();

        }


        public StringBuilder SETFASETSQL1(string QUERYS)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();        

            FASTSQL.AppendFormat(@"      
                                SELECT 
                                TG001 AS '進貨單別'
                                ,TG002 AS '進貨單號'
                                ,TG003 AS '進貨日'
                                ,TG021 AS '廠商'
                                ,TH004 AS '品號'
                                ,TH005 AS '品名'
                                ,TH006 AS '規格'
                                ,TH007 AS '數量'
                                ,TH008 AS '單位'
                                ,TG031+TG032 AS '總金額'
                                ,TH047+TH048 AS '明細金額'
                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH
                                WHERE TG001=TH001 AND TG002=TH002 
                                AND TG013 IN ('Y')
                                AND TG033 IN (
	                                SELECT  [TG033]
	                                FROM [TKPUR].[dbo].[TKPURTGTG033]
                                )
                                AND TG001+TG002 IN ({0})
                                ORDER BY TG001,TG002,TG003
                               
                                ", QUERYS.ToString());

            return FASTSQL;
        }

        public void Search2(string TC002, string MA001)
        {
            DataSet ds = new DataSet();

            StringBuilder sbSqlQuery1 = new StringBuilder();
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

                if (!string.IsNullOrEmpty(TC002))
                {
                    sbSqlQuery1.AppendFormat(@" 
                                            AND 單號 LIKE '%{0}%'
                                                ", TC002);
                }
                else
                {
                    sbSqlQuery1.AppendFormat(@" 
                                           
                                                ");
                }


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



                //採購的進貨+製令的託外進貨
                sbSql.AppendFormat(@"                                    
                                   SELECT *
                                    FROM 
                                    (
                                    SELECT 
                                    TC001 AS '單別'
                                    ,TC002 AS '單號'
                                    ,TC003 AS '採購日期'                                    
                                    ,MA002 AS '廠商全名'
                                    ,(CASE WHEN TC018=1 THEN '應稅內含' 
                                    WHEN TC018=2 THEN '應稅外加' 
                                    WHEN TC018=3 THEN '零稅率' 
                                    WHEN TC018=4 THEN '免稅' 
                                    WHEN TC018=9 THEN '不計稅' 
                                    END) AS '課稅別'
                                    ,TC019 AS '本幣貨款金額'
                                    ,TC020 AS '本幣稅額'
                                    ,(TC019+TC020) AS '本幣合計金額'
                                    ,TC004 AS '供應廠商'
                                    ,TC027
                                    FROM [TK].dbo.PURTC
                                    LEFT JOIN [TK].dbo.PURMA ON MA001=TC004
                                    WHERE 1=1
                                    ) AS TEMP
                                    WHERE 1=1   
                                    AND TC027 IN (
	                                    SELECT  [TG033]
	                                    FROM [TKPUR].[dbo].[TKPURTGTG033]
                                    )                                
                                    {0}
                                    {1}
                                    ORDER BY 單別,單號
                               
                                    ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;

                    MessageBox.Show("查無資料");
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["TEMPds1"];
                        dataGridView2.AutoResizeColumns();

                        dataGridView2.AutoResizeColumns();
                        dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                        dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                        // 設定數字格式
                        // 或使用 "N2" 表示兩位小數點（例如：12,345.67）
                        dataGridView2.Columns["本幣貨款金額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點
                        dataGridView2.Columns["本幣稅額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點
                        dataGridView2.Columns["本幣合計金額"].DefaultCellStyle.Format = "N0"; // 每三位一個逗號，無小數點



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
        /// <summary>
        /// 找出 DataGridView 中被勾選的行的名稱。
        /// </summary>
        /// <param name="dataGridView">要操作的 DataGridView 控制項。</param>
        /// <param name="checkBoxColumnName">CheckBox 欄位的 Name 屬性值 (例如: "SelectedCheckbox")。</param>
        /// <param name="nameColumnName">包含名稱資料的欄位 Name 屬性值 (例如: "Name")。</param>
        /// <returns>一個包含所有被勾選名稱的 List<string>。</returns>
        public List<string> GetSelectedNamesFromDGV2(DataGridView dataGridView, string checkBoxColumnName)
        {
            List<string> selectedNames = new List<string>();

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                // 忽略新增行 (New Row)
                if (row.IsNewRow)
                {
                    continue;
                }

                // 確保兩個目標欄位都存在
                if (row.Cells[checkBoxColumnName] != null && row.Cells["單號"] != null)
                {
                    // 獲取 CheckBox 欄位的值
                    object checkBoxValue = row.Cells[checkBoxColumnName].Value;
                    bool isChecked = false;

                    // 嘗試將值轉換為 bool。需要處理 null 或 DBNull 的情況。
                    if (checkBoxValue != null && checkBoxValue != DBNull.Value)
                    {
                        try
                        {
                            // 2. 使用 Convert.ToBoolean 進行轉換
                            // 它能處理 bool, 1/0 (int), "True"/"False" (string) 等情況
                            isChecked = Convert.ToBoolean(checkBoxValue);
                        }
                        catch (InvalidCastException)
                        {
                            // 處理轉換失敗的情況 (例如欄位包含非 1/0 的文字)
                            // 您可以記錄錯誤或設定預設值
                            isChecked = false;
                        }
                        catch (FormatException)
                        {
                            // 處理格式錯誤 (例如欄位是無效字串)
                            isChecked = false;
                        }
                    }

                    // 如果 CheckBox 被勾選，則獲取對應的名稱
                    if (isChecked)
                    {
                        // 獲取名稱欄位的值 (通常是 string)
                        object TC001 = row.Cells["單別"].Value;
                        object TC002 = row.Cells["單號"].Value;
                        if (TC002 != null && TC002 != DBNull.Value)
                        {
                            selectedNames.Add(TC001.ToString().Trim() + TC002.ToString().Trim());
                        }
                    }
                }
            }

            return selectedNames; 
        } 

        public void SETFASTREPORT2(string sqlInCondition)
        {
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();
            report1.Load(@"REPORT\請款憑單BY採購單.frx");

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

            SQL = SETFASETSQL2(sqlInCondition);

            Table.SelectCommand = SQL.ToString(); ;

            //report1.SetParameterValue("P1", COMMENT);

            report1.Preview = previewControl2;
            report1.Show();

        }


        public StringBuilder SETFASETSQL2(string QUERYS)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();

            FASTSQL.AppendFormat(@"  
                                SELECT 
                                TC001 AS '採購單別'
                                ,TC002 AS '採購單號'
                                ,TC003 AS '採購日'
                                ,MA002 AS '廠商'
                                ,TD004 AS '品號'
                                ,TD005 AS '品名'
                                ,TD006 AS '規格'
                                ,TD008 AS '數量'
                                ,TD009 AS '單位'
                                ,TC019+TC020 AS '總金額'
                                ,CASE WHEN TC018='2' THEN ROUND(TD011*(1+TC026),0) ELSE TD011 END AS '明細金額'
                                ,TC027
                                FROM [TK].dbo.PURTC
                                LEFT JOIN [TK].dbo.PURMA ON MA001=TC004
                                ,[TK].dbo.PURTD
                                WHERE TC001=TD001 AND TC002=TD002 
                                AND TC014 IN ('Y')
                                AND TC027 IN (
	                                SELECT  [TG033]
	                                FROM [TKPUR].[dbo].[TKPURTGTG033]
                                )
                                AND TC001+TC002 IN ({0})
                                ORDER BY TC001,TC002,TC003
                               
                                ", QUERYS.ToString());

            return FASTSQL;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox5.Text.Trim(), textBox6.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            string sqlInCondition = "";
            // 假設您的 DataGridView 叫 dataGridView1
            // 假設 CheckBox 欄位的 Name 是 "SelectedCheckbox"
            // 假設名稱欄位的 Name 是 "Name"
            List<string> names = GetSelectedNamesFromDGV(dataGridView1, "SelectedCheckbox");

            if (names.Count > 0)
            {
                // 關鍵步驟：為每個名稱加上單引號，同時處理名稱中可能包含的單引號
                // 將每個單引號 ' 替換成兩個單引號 '' (這是 SQL 轉義字元)
                var escapedNames = names.Select(name =>
                    $"'{name.Replace("'", "''")}'"
                );

                // 使用逗號將所有格式化後的字串連接起來
                sqlInCondition = string.Join(", ", escapedNames);
            }
            else
            {
                
            }

            SETFASTREPORT(sqlInCondition);
            //MessageBox.Show(sqlInCondition.ToString());
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Search2(textBox1.Text.Trim(), textBox2.Text.Trim());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sqlInCondition = "";
            // 假設您的 DataGridView 叫 dataGridView1
            // 假設 CheckBox 欄位的 Name 是 "SelectedCheckbox"
            // 假設名稱欄位的 Name 是 "Name"
            List<string> names = GetSelectedNamesFromDGV2(dataGridView2, "SelectedCheckbox");

            if (names.Count > 0)
            {
                // 關鍵步驟：為每個名稱加上單引號，同時處理名稱中可能包含的單引號
                // 將每個單引號 ' 替換成兩個單引號 '' (這是 SQL 轉義字元)
                var escapedNames = names.Select(name =>
                    $"'{name.Replace("'", "''")}'"
                );

                // 使用逗號將所有格式化後的字串連接起來
                sqlInCondition = string.Join(", ", escapedNames);
            }
            else
            {

            }

            SETFASTREPORT2(sqlInCondition);
            //MessageBox.Show(sqlInCondition.ToString());
        }

        #endregion


    }
}
