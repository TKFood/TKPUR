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

                                    FROM [TK].dbo.PURTG
                                    WHERE 1=1

                                    UNION ALL
                                    SELECT 
                                    TH001 AS '單別'
                                    ,TH002 AS '單號'
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
                                    ,TH005 AS '供應廠商'

                                    FROM [TK].dbo.MOCTH,[TK].dbo.PURMA
                                    WHERE TH005=MA001
                                    ) AS TEMP
                                    WHERE 1=1                                   
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
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search( textBox5.Text.Trim(), textBox6.Text.Trim());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            // 假設您的 DataGridView 叫 dataGridView1
            // 假設 CheckBox 欄位的 Name 是 "SelectedCheckbox"
            // 假設名稱欄位的 Name 是 "Name"
            List<string> names = GetSelectedNamesFromDGV(dataGridView1, "SelectedCheckbox");

            if (names.Count > 0)
            {
                MessageBox.Show($"被勾選的名稱有：\n{string.Join(", ", names)}");
            }
            else
            {
                MessageBox.Show("沒有任何項目被勾選。");
            }
        }
        #endregion


    }
}
