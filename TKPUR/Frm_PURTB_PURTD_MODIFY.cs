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
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKPUR
{
    public partial class Frm_PURTB_PURTD_MODIFY : Form
    {
        public Frm_PURTB_PURTD_MODIFY()
        {
            InitializeComponent();
        }
        private void Frm_PURTB_PURTD_MODIFY_Load(object sender, EventArgs e)
        {
            SET_comboBox1();
        }


        #region FUNCTION
        public void SET_comboBox1()
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Y");
            comboBox1.Items.Add("N");
        }

        public void SEARCH_DG1(string SDATES)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    TA001 AS '單別'
                                    ,TA002 AS '單號'
                                    ,TB003 AS '序號'
                                    ,TB004 AS '品號'
                                    ,TB005 AS '品名'
                                    ,TB039  AS '結案碼'
                                    FROM [TK].dbo.PURTB
                                    INNER JOIN [TK].dbo.PURTA ON TA001=TB001 AND TA002=TB002
                                    WHERE TA002 LIKE '%{0}%'
                                    ORDER BY TA001,TA002,TB003

                                    ", SDATES);
                SqlDataAdapter da = new SqlDataAdapter(@"" + sbSql, sqlConn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;


            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX_NULL();

            if (dataGridView1.CurrentRow != null)
            {
                string ta001 = dataGridView1.CurrentRow.Cells["單別"].Value.ToString();
                string ta002 = dataGridView1.CurrentRow.Cells["單號"].Value.ToString();
                string tb003 = dataGridView1.CurrentRow.Cells["序號"].Value.ToString();
                string tb039 = dataGridView1.CurrentRow.Cells["結案碼"].Value.ToString();

                textBox1.Text = ta001;
                textBox2.Text = ta002;
                textBox3.Text = tb003;
                comboBox1.Text = tb039;
            }

        }

        public void UPDATE_PURTB(string TA001, string TA002, string TB003, string TB039)
        {
            string sql = @"
                            UPDATE [TK].dbo.PURTB 
                            SET TB039 = @TB039 
                            WHERE TB001 = @TB001 
                              AND TB002 = @TB002 
                              AND TB003 = @TB003";

            try
            {
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                // 使用 using 自動管理 SqlConnection 與 SqlCommand 的開啟與釋放
                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(sql, sqlConn))
                    {
                        // 綁定 SQL 參數 (將傳入的 TA001/TA002 對應至 SQL 的 TB001/TB002)
                        cmd.Parameters.AddWithValue("@TB001", (object)TA001 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TB002", (object)TA002 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TB003", (object)TB003 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TB039", (object)TB039 ?? DBNull.Value);

                        sqlConn.Open();          // 1. 開啟連線
                        cmd.ExecuteNonQuery();   // 2. 使用正確的同步執行方法
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("UPDATE_PURTB 執行失敗: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// 搜尋並恢復 DataGridView 的選取列與捲軸位置
        /// </summary>
        private void RestoreGridPosition_DG1(string ta001, string ta002,string tb003, int originalScrollIndex)
        {
            if (string.IsNullOrEmpty(ta001) || string.IsNullOrEmpty(ta002)) return;

            bool isFound = false;

            // 逐列搜尋剛剛修改的主鍵資料
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells["單別"].Value?.ToString() == ta001 &&
                    row.Cells["單號"].Value?.ToString() == ta002 &&
                    row.Cells["序號"].Value?.ToString() == tb003)
                {
                    dataGridView1.ClearSelection(); // 清除先前的選取狀態

                    // 設定游標焦點至該列的第一個可見儲存格
                    dataGridView1.CurrentCell = row.Cells[0];
                    row.Selected = true;

                    isFound = true;
                    break;
                }
            }

            // 恢復捲軸位置
            if (isFound && originalScrollIndex >= 0 && originalScrollIndex < dataGridView1.RowCount)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = originalScrollIndex;
            }
        }

        public void SET_TEXTBOX_NULL()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");            
            SEARCH_DG1(SDATES);
        }


        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            // 1. 檢查目前是否有選取資料列
            if (dataGridView1.CurrentRow == null) return;

            string ta001=textBox1.Text; 
            string ta002 = textBox2.Text;
            string tb003 = textBox3.Text;
            string tb039 = comboBox1.Text;
            UPDATE_PURTB(ta001, ta002, tb003, tb039);

            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            SEARCH_DG1(SDATES);

            int scrollIndex = dataGridView1.FirstDisplayedScrollingRowIndex;
            // 5. 將游標與捲軸定位回到剛才那筆資料
            RestoreGridPosition_DG1((ta001, ta002, tb003, scrollIndex);
        }
    }
}
