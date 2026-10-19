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
        //時間單位：秒（Seconds），設定 SQL 最多 4 分鐘
        private const int SQL_TIMEOUT_LIMITS = 60 * 4;
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapterMAIL = new SqlDataAdapter();

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
                                TD012 AS '預交日',
                                TC001 AS '採購單別',
                                TC002 AS '採購單號',
                                TC004 AS '供應廠商',
                                MA002 AS '廠商',
                                TD003 AS '採購序號',
                                TD004 AS '品號',
                                TD005 AS '品名',
                                TD008 AS '採購數量',
                                TD015 AS '已交數量',
                                TD016 AS '結案碼',
                                TD013 AS '訂單單別',
                                TD021 AS '訂單單號',
                                TD023 AS '訂單序號'

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
                    string td003 = row.Cells["採購序號"].Value?.ToString() ?? "";
                    string TD013 = row.Cells["訂單單別"].Value?.ToString() ?? "";
                    string TD021 = row.Cells["訂單單號"].Value?.ToString() ?? "";
                    string TD023 = row.Cells["訂單序號"].Value?.ToString() ?? "";
                    // 組合成 "TC001,TC002,TD003" 格式字串
                    string selectedRow = $"{tc001},{tc002},{td003},{TD013},{TD021},{TD023}";
                    selectedRows.Add(selectedRow);
                }
            }

            return selectedRows;
        }
        // 2. 將 LIST 中的資料拆解並更新到資料庫中
        public void UPDATE_TD016(string TD016, List<string> selectedRows)
        {
            //List<string> selectedRows = GetSelectedRows();

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
                                if (parts.Length == 6)
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

        //發送MAIL
        public void SEND_MAIL(List<string> selectedRows,string TD016)
        {
            DataTable DS_EMAIL_TO_EMAIL = new DataTable();
            DataTable DT_DATAS = new DataTable();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            try
            {
                DS_EMAIL_TO_EMAIL = SERACH_MAIL_SEND_TO();
                //DT_DATAS = SERACH_DEC_NEW_PRODUCT_PRICES_UOF(cancellationToken);
                
                SUBJEST.Clear();
                BODY.Clear();


                SUBJEST.AppendFormat(@"系統通知-請查收-採購-大潁的採購單(含訂單)的明細，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                //ERP 採購相關單別、單號未核準的明細
                //
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                    + "<br>" + "系統通知-請查收-採購-大潁的採購單(含訂單)的明細，謝謝"
                    + " <br>"
                    );


                if (TD016.Equals("Y"))
                {
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>"
                                       + "<br>" + "以下是 結案的 大潁的採購單(含訂單)的明細，謝謝"
                                       + " <br>"
                                       );
                }
                else 
                {
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>"
                                          + "<br>" + "以下是 未結案 的大潁的採購單(含訂單)的明細，謝謝"
                                          + " <br>"
                                          );
                }

                if (selectedRows != null && selectedRows.Count >= 1)
                {
                    if (selectedRows.Count > 0)
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                        BODY.AppendFormat(@"<table> ");
                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單別</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購序號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">訂單單別</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">訂單單號</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">訂單序號</th>");

                        BODY.AppendFormat(@"</tr> ");

                        foreach (string data in selectedRows)
                        {
                            // 依照逗號拆解出 TC001, TC002, TD003
                            string[] parts = data.Split(',');
                            if (parts.Length == 6)
                            {
                                string TC001 = parts[0];                               
                                string TC002 = parts[1];
                                string TD003 = parts[2];
                                string TD013 = parts[3];
                                string TD021 = parts[4];
                                string TD023 = parts[5];

                                BODY.AppendFormat(@"<tr >");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + TC001 + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + TC002 + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + TD003 + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + TD013 + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + TD021 + "</td>");
                                BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + TD023 + "</td>");

                                BODY.AppendFormat(@"</tr> ");

                            }
                        }
                        BODY.AppendFormat(@"</table> ");
                    }
                }
                else
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "無資料");
                }

                try
                {
                    string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
                    string NAME = ConfigurationManager.AppSettings["NAME"];
                    string PW = ConfigurationManager.AppSettings["PW"];

                    System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
                    MyMail.From = new System.Net.Mail.MailAddress("tk290@tkfood.com.tw");

                    //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
                    //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
                    MyMail.Subject = SUBJEST.ToString();
                    //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
                    MyMail.Body = BODY.ToString();
                    MyMail.IsBodyHtml = true; //是否使用html格式

                    //加上附圖
                    //string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                    //MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

                    System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
                    MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);


                    try
                    {
                        //通知預設的群組
                        foreach (DataRow DR in DS_EMAIL_TO_EMAIL.Rows)
                        {
                            MyMail.To.Add(DR["MAIL"].ToString()); //設定收件者Email，多筆mail
                        }

                        //通知未簽核人員 
                        //HashSet<string> 會自動過濾重複的 Email，確保 To 清單中不會有重複地址。
                        HashSet<string> emailSet = new HashSet<string>();
                        foreach (DataRow DR in DT_DATAS.Rows)
                        {
                            string email = DR["EMAIL"].ToString();
                            string SEICALNAMES = DR["NAME"].ToString();

                            if (!string.IsNullOrWhiteSpace(email) && emailSet.Add(email)) // 確保唯一
                            {
                                MyMail.To.Add(email);
                            }

                        }

                        //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源

                    }
                    catch (Exception EX)
                    {
                        //MessageBox.Show("有錯誤");
                    }
                }
                catch (Exception EX)
                {

                }
                finally
                {

                }
            }
            catch (Exception EX)
            {

            }
            finally
            {

            }
        }

        public DataTable SERACH_MAIL_SEND_TO()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
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

                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                   SELECT
                                [KINDS]
                                ,[KINDNAMES]
                                ,[MAIL]
                                ,[NAMES]
                                FROM [TKPUR].[dbo].[TB_PUR_MAILS_TO]
                                WHERE [KINDS]='DYPURTC'

                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = SQL_TIMEOUT_LIMITS;
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
            catch (Exception EX)
            {
                return null;
            }
            finally
            {

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



        private void button2_Click(object sender, EventArgs e)
        {
            List<string> selectedRows = GetSelectedRows();
            //修改成結案+發MAIL
            UPDATE_TD016("Y", selectedRows);

            SEND_MAIL(selectedRows,"Y");

            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            string TC002 = textBox1.Text.ToString();
            string TD016 = comboBox1.Text.ToString();
            SearchData(SDATE, EDATE, TC002, TD016);

            MessageBox.Show("完成");

        }

        private void button3_Click(object sender, EventArgs e)
        {
            List<string> selectedRows = GetSelectedRows();
            //修改成未結案+發MAIL
            UPDATE_TD016("N", selectedRows);

            SEND_MAIL(selectedRows, "N");

            string SDATE = dateTimePicker1.Value.ToString("yyyyMMdd");
            string EDATE = dateTimePicker2.Value.ToString("yyyyMMdd");
            string TC002 = textBox1.Text.ToString();
            string TD016 = comboBox1.Text.ToString();
            SearchData(SDATE, EDATE, TC002, TD016);

            MessageBox.Show("完成");
        }

        #endregion
    }
}
