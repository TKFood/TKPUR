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
            SET_comboBox2();    
        }


        #region FUNCTION
        public void SET_comboBox1()
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Y");
            comboBox1.Items.Add("N");
        }
        public void SET_comboBox2()
        {
            comboBox2.Items.Clear();
            comboBox2.Items.Add("Y");
            comboBox2.Items.Add("N");
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

        public void SEARCH_DG2(string SDATES)
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
                                    TC001 AS '單別'
                                    ,TC002 AS '單號'
                                    ,TD003 AS '序號'
                                    ,TD004 AS '品號'
                                    ,TD005 AS '品名'
                                    ,TD016  AS '結案碼'
                                    ,TD026 AS '請購單別'
                                    ,TD027 AS '請購單號'
                                    ,TD028 AS '請購序號'
                                    FROM [TK].dbo.PURTD
                                    INNER JOIN [TK].dbo.PURTC ON TC001=TD001 AND TC002=TD002
                                    WHERE TC002 LIKE '%{0}%'
                                    ORDER BY TC001,TC002,TC003

                                    ", SDATES);
                SqlDataAdapter da = new SqlDataAdapter(@"" + sbSql, sqlConn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView2.DataSource = dt;


            }
            catch
            {

            }
            finally
            {

            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX_NULL_DG2();

            if (dataGridView2.CurrentRow != null)
            {
                string tc001 = dataGridView2.CurrentRow.Cells["單別"].Value.ToString();
                string tc002 = dataGridView2.CurrentRow.Cells["單號"].Value.ToString();
                string td003 = dataGridView2.CurrentRow.Cells["序號"].Value.ToString();
                string td016 = dataGridView2.CurrentRow.Cells["結案碼"].Value.ToString();
                string td026 = dataGridView2.CurrentRow.Cells["請購單別"].Value.ToString();
                string td027 = dataGridView2.CurrentRow.Cells["請購單號"].Value.ToString();
                string td028 = dataGridView2.CurrentRow.Cells["請購序號"].Value.ToString();

                textBox4.Text = tc001;
                textBox5.Text = tc002;
                textBox6.Text = td003;
                comboBox2.Text = td016;

                textBox7.Text = td026;
                textBox8.Text = td027;
                textBox9.Text = td028;

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


        public void UPDATE_PURTD(string TC001, string TC002, string TD003, string TD016, string TD026,string TD027, string TD028)
        {
            string sql = @"
                            UPDATE [TK].dbo.PURTD 
                            SET TD016 = @TD016 ,
                                TD026=@TD026,
                                TD027 = @TD027,
                                TD028 = @TD028
                            WHERE TD001 = @TD001     
                              AND TD002 = @TD002
                              AND TD003 = @TD003";

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
                        cmd.Parameters.AddWithValue("@TD001", (object)TC001 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TD002", (object)TC002 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TD003", (object)TD003 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TD016", (object)TD016 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TD026", (object)TD026 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TD027", (object)TD027 ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@TD028", (object)TD028 ?? DBNull.Value);

                        sqlConn.Open();          // 1. 開啟連線
                        cmd.ExecuteNonQuery();   // 2. 使用正確的同步執行方法
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("UPDATE_PURTD 執行失敗: " + ex.Message, ex);
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

        /// <summary>
        /// 搜尋並恢復 DataGridView 的選取列與捲軸位置
        /// </summary>
        private void RestoreGridPosition_DG2(string tc001, string tc002, string td003, int originalScrollIndex)
        {
            if (string.IsNullOrEmpty(tc001) || string.IsNullOrEmpty(tc002)) return;

            bool isFound = false;

            // 逐列搜尋剛剛修改的主鍵資料
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Cells["單別"].Value?.ToString() == tc001 &&
                    row.Cells["單號"].Value?.ToString() == tc002 &&
                    row.Cells["序號"].Value?.ToString() == td003)
                {
                    dataGridView2.ClearSelection(); // 清除先前的選取狀態

                    // 設定游標焦點至該列的第一個可見儲存格
                    dataGridView2.CurrentCell = row.Cells[0];
                    row.Selected = true;

                    isFound = true;
                    break;
                }
            }

            // 恢復捲軸位置
            if (isFound && originalScrollIndex >= 0 && originalScrollIndex < dataGridView2.RowCount)
            {
                dataGridView2.FirstDisplayedScrollingRowIndex = originalScrollIndex;
            }
        }

        public void SET_TEXTBOX_NULL()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            
        }
        public void SET_TEXTBOX_NULL_DG2()
        {
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");            
            SEARCH_DG1(SDATES);
        }


        private void button3_Click(object sender, EventArgs e)
        {
            // 1. 檢查目前是否有選取資料列
            if (dataGridView1.CurrentRow == null) return;

            string ta001 = textBox1.Text;
            string ta002 = textBox2.Text;
            string tb003 = textBox3.Text;
            string tb039 = comboBox1.Text;
            UPDATE_PURTB(ta001, ta002, tb003, tb039);

            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");
            SEARCH_DG1(SDATES);

            int scrollIndex = dataGridView1.FirstDisplayedScrollingRowIndex;
            // 5. 將游標與捲軸定位回到剛才那筆資料
            RestoreGridPosition_DG1(ta001, ta002, tb003, scrollIndex);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            SEARCH_DG2(SDATES);
        }
        #endregion

        private void button6_Click(object sender, EventArgs e)
        {
            // 1. 檢查目前是否有選取資料列
            if (dataGridView2.CurrentRow == null) return;

            string tc001 = textBox4.Text;
            string tc002 = textBox5.Text;
            string tc003 = textBox6.Text;
            string td016 = comboBox2.Text;
            string td026 = textBox7.Text;
            string td027 = textBox8.Text;
            string td028 = textBox9.Text;

            UPDATE_PURTD(tc001, tc002, tc003, td016, td026, td027, td028);

            string SDATES = dateTimePicker2.Value.ToString("yyyyMMdd");
            SEARCH_DG2(SDATES);

            int scrollIndex = dataGridView2.FirstDisplayedScrollingRowIndex;
            // 5. 將游標與捲軸定位回到剛才那筆資料
            RestoreGridPosition_DG2(tc001, tc002, tc003, scrollIndex);
        }
    }
}
