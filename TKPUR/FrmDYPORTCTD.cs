using FastReport;
using FastReport.Data;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TKITDLL;

namespace TKPUR
{
    public partial class FrmDYPORTCTD : Form
    {
        public FrmDYPORTCTD()
        {
            InitializeComponent();
        }
        private void FrmDYPORTCTD_Load(object sender, EventArgs e)
        {
            SETDATES();
            AddComboBoxItems();
            AddCheckBoxColumn();
        }
        #region FUNCTION
        //在DATAGRIDVIEW前，加上勾選欄位
        private void AddCheckBoxColumn()
        {
            DataGridViewCheckBoxColumn checkBoxColumn = new DataGridViewCheckBoxColumn();
            checkBoxColumn.HeaderText = "選擇";
            checkBoxColumn.Width = 50;
            checkBoxColumn.Name = "checkBoxColumn";
            dataGridView1.Columns.Insert(0, checkBoxColumn);
        }
        //增加COMBOX1的ITESM
        private void AddComboBoxItems()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                StringBuilder sb = new StringBuilder();
                StringBuilder sb1 = new StringBuilder();

               


                sb.AppendFormat(@"
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='FrmDYPORTCTD'
                                ORDER BY [ID]
                                ");

                SqlCommand command = new SqlCommand(sb.ToString(), connection);
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataSet ds = new DataSet();
                da.Fill(ds, "TEMPds1");
                
                comboBox1.Items.Clear();
                foreach (DataRow dr in ds.Tables["TEMPds1"].Rows)
                {
                    comboBox1.Items.Add(dr["PARANAME"].ToString());
                }
                comboBox1.Text = ds.Tables["TEMPds1"].Rows[0]["PARANAME"].ToString();

            }
        }
        //SETDATES
        public void SETDATES()
        {
            DateTime now = DateTime.Now;
            int year = now.Year;
            int month = now.Month;
            int day = now.Day;
            // 設定起始日期為當月的第一天
            DateTime startDate = new DateTime(year, 1, 1);
            dateTimePicker1.Value = startDate;
            // 設定結束日期為當月的最後一天
            DateTime endDate = new DateTime(year, month, DateTime.DaysInMonth(year, month));
            dateTimePicker2.Value = endDate;
        }
        //查詢資料表，顯示到dataGridView1
        public void SearchData(string SDATE, string EDATE,string TC002,string TD016)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                StringBuilder sb = new StringBuilder();
                StringBuilder sb1 = new StringBuilder();
                StringBuilder sb2 = new StringBuilder();

                if (!string.IsNullOrEmpty(TC002))
                {
                    sb1.AppendFormat(" AND TC002 LIKE '%{0}%' ", TC002);
                }
                else
                {
                    sb1.AppendFormat("  ");
                }              
                if (!string.IsNullOrEmpty(TD016)&& TD016.Equals("Y"))
                {
                    sb2.AppendFormat(" AND TD016 IN ('Y','y')");
                }
                else if (!string.IsNullOrEmpty(TD016) && TD016.Equals("N"))
                {
                    sb2.AppendFormat(" AND TD016 NOT IN ('Y','y') ");
                }
                else 
                {
                    sb2.AppendFormat(" ");
                }


                sb.AppendFormat(@"
                                SELECT 
                                TC001 AS '採購單別',
                                TC002 AS '採購單號',
                                TC004 AS '供應廠商',
                                MA002 AS '廠商',
                                TD003 AS '序號',
                                TD004 AS '品號',
                                TD005 AS '品名',
                                TD008 AS '採購數量',
                                TD015 AS '已交數量',
                                TD016 AS '結案碼'

                                FROM [DY].dbo.PURTD,[DY].dbo.PURTC
                                LEFT JOIN [DY].dbo.PURMA ON MA001=TC004
                                WHERE TC001=TD001 AND TC002=TD002                                
                                AND TC002>='{0}' AND TC002<='{1}'
                                {2}
                                {3}

                                ", SDATE, EDATE, sb1.ToString(), sb2.ToString());

                SqlCommand command = new SqlCommand(sb.ToString(), connection);
                SqlDataAdapter da = new SqlDataAdapter(command);
                DataSet ds = new DataSet();
                da.Fill(ds, "TEMPds1");
                dataGridView1.DataSource = ds.Tables["TEMPds1"];
            }
        }
        public List<string> GetSelectedRows()
        {
            List<string> selectedRows = new List<string>();

            // 確保當前編輯中的儲存格資料已寫入 DataGridView
            dataGridView1.EndEdit();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // 略過新增用的空白列
                if (row.IsNewRow) continue;

                DataGridViewCheckBoxCell checkBoxCell = row.Cells["checkBoxColumn"] as DataGridViewCheckBoxCell;
                if (checkBoxCell != null && checkBoxCell.Value != null && Convert.ToBoolean(checkBoxCell.Value))
                {
                    string tc001 = row.Cells["採購單別"].Value?.ToString() ?? "";
                    string tc002 = row.Cells["採購單號"].Value?.ToString() ?? "";
                    string td003 = row.Cells["序號"].Value?.ToString() ?? "";

                    // 組合成 "TC001,TC002,TD003" 格式字串
                    string selectedRow = $"{tc001},{tc002},{td003}";
                    selectedRows.Add(selectedRow);
                }
            }

            return selectedRows;
        }
        // 2. 將 LIST 中的資料拆解並更新到資料庫中
        public void UPDATE_TD016(string TD016)
        {
            List<string> selectedRows = GetSelectedRows();

            if (selectedRows == null || selectedRows.Count == 0)
            {
                MessageBox.Show("請先勾選要更新的資料！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Class1 TKID = new Class1();
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            // 資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                connection.Open();

                // 使用 SqlTransaction 確保批次更新的一致性（全部成功或全部失敗）
                using (SqlTransaction transaction = connection.BeginTransaction())
                {
                    try
                    {
                        // 假設您的目標資料表是 PURTD，要更新 TD016 欄位
                        string updateSql = @"
                                            UPDATE [DY].dbo.PURTD 
                                            SET TD016 = @TD016 
                                            WHERE TD001 = @TC001 
                                              AND TD002 = @TC002 
                                              AND TD003 = @TD003";

                        using (SqlCommand cmd = new SqlCommand(updateSql, connection, transaction))
                        {
                            // 預先宣告參數型態，提升迴圈執行效能
                            cmd.Parameters.Add("@TD016", SqlDbType.VarChar);
                            cmd.Parameters.Add("@TC001", SqlDbType.VarChar);
                            cmd.Parameters.Add("@TC002", SqlDbType.VarChar);
                            cmd.Parameters.Add("@TD003", SqlDbType.VarChar);

                            foreach (string data in selectedRows)
                            {
                                // 依照逗號拆解出 TC001, TC002, TD003
                                string[] parts = data.Split(',');
                                if (parts.Length == 3)
                                {
                                    cmd.Parameters["@TD016"].Value = TD016; // 依需求設定要更新的值 (例如 'Y')
                                    cmd.Parameters["@TC001"].Value = parts[0];
                                    cmd.Parameters["@TC002"].Value = parts[1];
                                    cmd.Parameters["@TD003"].Value = parts[2];

                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        // 提交交易
                        transaction.Commit();
                        MessageBox.Show("更新成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        // 發生例外時全部復原，避免部分更新
                        transaction.Rollback();
                        MessageBox.Show($"更新資料庫時發生錯誤：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            string TC002= textBox1.Text.ToString();
            string TD016 = comboBox1.Text.ToString();
            SearchData(SDATE, EDATE, TC002, TD016);   
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            //修改成結案+發MAIL
            UPDATE_TD016("Y");

            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            string TC002 = textBox1.Text.ToString();
            string TD016 = comboBox1.Text.ToString();
            SearchData(SDATE, EDATE, TC002, TD016);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //修改成未結案+發MAIL
            UPDATE_TD016("N");

            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            string TC002 = textBox1.Text.ToString();
            string TD016 = comboBox1.Text.ToString();
            SearchData(SDATE, EDATE, TC002, TD016);
        }
    }
}
