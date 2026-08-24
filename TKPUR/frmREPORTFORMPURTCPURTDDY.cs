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
    public partial class frmREPORTFORMPURTCPURTDDY : Form
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

        DataSet DSPURTCTD = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        Report report1 = new Report();
        string MAILATTACHPATH = null;
        DataSet DSMAIL = new DataSet();

        string PDF_PATH = "";

        public frmREPORTFORMPURTCPURTDDY()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION
        private void frmREPORTFORMPURTCPURTD_Load(object sender, EventArgs e)
        {
            SETGRIDVIEW();

            SET_PDF_PATH();
        }

        public string SET_PDF_PATH()
        {
            PDF_PATH = @"\\192.168.1.109\部門檔案區\1300 資材部群組\傳真記錄區-待寄";

            return PDF_PATH;
        }
        public void SETGRIDVIEW()
        {
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色

            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "　勾選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView1.Columns.Insert(0, cbCol);

            #region 建立全选 CheckBox

            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView1.Controls.Add(cbHeader);

            #endregion
        }

        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView1.EndEdit();

            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView1.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT FORM FROM [TKPUR].[dbo].[PURREPORTFORM] WHERE [REPORT]='憑証回傳' ORDER BY ID  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("FORM", typeof(string));

            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "FORM";
            comboBox1.DisplayMember = "FORM";
            sqlConn.Close();


        }

        public void Search(string SDAY, string EDAY)
        {
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
                                        ISNULL(TD_Agg.DetailCount, 0) AS [明細筆數],
                                        TC.TC001 AS [採購單別],
                                        TC.TC002 AS [採購單號],
                                        TC.TC003 AS [採購日期],
                                        TC.TC004 AS [供應廠商],
                                        MA.MA002 AS [供應廠],
                                        MA.MA011 AS [EMAIL],
                                        ISNULL(TD_Agg.Details, '')   AS [明細],
                                        UOF.[COMMENT]                AS [採購簽核意見]
                                    FROM [DY].[dbo].[PURTC] TC WITH (NOLOCK)
                                    INNER JOIN [DY].[dbo].[PURMA] MA WITH (NOLOCK)
                                        ON TC.TC004 = MA.MA001

                                    -- 1. 使用 OUTER APPLY 一次性處理 PURTD（明細筆數 + 字串合併，且去除結尾多餘逗號）
                                    OUTER APPLY (
                                        SELECT 
                                            COUNT(TD.TD004) AS DetailCount,
                                            STUFF((
                                                SELECT ', ' + ISNULL(TD004,'') + ISNULL(TD005,'') + ISNULL(TD006,'')
                                                FROM [DY].[dbo].[PURTD] TD WITH (NOLOCK)
                                                WHERE TD.TD001 = TC.TC001 
                                                  AND TD.TD002 = TC.TC002
                                                FOR XML PATH(''), TYPE
                                            ).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS Details
                                        FROM [DY].[dbo].[PURTD] TD WITH (NOLOCK)
                                        WHERE TD.TD001 = TC.TC001 
                                          AND TD.TD002 = TC.TC002
                                    ) AS TD_Agg

                                    -- 2. 外部 Linked Server 採用 OUTER APPLY 取 TOP 1（效能比 SELECT 子查詢好）
                                    OUTER APPLY (
                                        SELECT TOP 1 [COMMENT]
                                        FROM [192.168.1.223].[UOF].[dbo].[TB_WKF_TASK_PUR_COMMENT] V
                                        WHERE V.[TC001] = TC.TC001 COLLATE Chinese_Taiwan_Stroke_BIN
                                          AND V.[TC002] = TC.TC002 COLLATE Chinese_Taiwan_Stroke_BIN
                                    ) AS UOF

                                    WHERE TC.TC003 >='{0}' AND TC.TC003<='{1}'
                                    ORDER BY TC.TC001, TC.TC002;
                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 將查詢逾時時間加長至 180 秒
                adapter.SelectCommand.CommandTimeout = 180;
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();


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
        public void PREPRINTS(string statusReports, string COMMENT)
        {
            string PRINTSPURTCPURTD = null;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(dr.Cells[0].Value) == true)
                {
                    //MessageBox.Show(dr.Cells["採購單別"].Value.ToString()+ dr.Cells["採購單號"].Value.ToString());

                    PRINTSPURTCPURTD = PRINTSPURTCPURTD + "'" + dr.Cells["採購單別"].Value.ToString().Trim() + dr.Cells["採購單號"].Value.ToString().Trim() + "',";
                }
            }

            PRINTSPURTCPURTD = PRINTSPURTCPURTD + "'A'";

            SETFASTREPORT(statusReports, PRINTSPURTCPURTD, COMMENT);
            //MessageBox.Show(PRINTSPURTCPURTD);
        }

        public void SETFASTREPORT(string statusReports, string PRINTSPURTCPURTD, string COMMENT)
        {
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();

            report1.Load(@"REPORT\採購單憑証V6-核準-雅芳-大潁.frx");
            //傳入SIGNS，當簽名檔的判斷
            report1.SetParameterValue("SIGNS", statusReports);

            //if (statusReports.Equals("憑証回傳"))
            //{
            //    report1.Load(@"REPORT\採購單憑証V6-無核準-大潁.frx"); 
            //}           
            //else if (statusReports.Equals("雅芳-簽名"))
            //{
            //    report1.Load(@"REPORT\採購單憑証V6-核準-雅芳-大潁.frx");
            //}
           
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
            // ★★★ 2. 加長 FastReport 查詢逾時時間（例如 180 秒；設為 0 代表無限制）★★★
            report1.Dictionary.Connections[0].CommandTimeout = 180;
            SQL = SETFASETSQL(statusReports, PRINTSPURTCPURTD);
       
            Table.SelectCommand = SQL.ToString(); 

            report1.SetParameterValue("P1", COMMENT);

            report1.Preview = previewControl1;
            report1.Show();

        }

        public StringBuilder SETFASETSQL(string statusReports, string PRINTSPURTCPURTD)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (statusReports.Equals("有簽名"))
            {
                STRQUERY.AppendFormat(@"  
                                        AND TC014 IN ('Y')
                                        ");
            }
            else
            {
                STRQUERY.AppendFormat(@"
                                        
                                        ");
            } 

            FASTSQL.AppendFormat(@"      
                                 SELECT *
                                ,CASE WHEN TC018='1' THEN '應稅內含' WHEN TC018='2' THEN '應稅外加' WHEN TC018='3' THEN '零稅率' WHEN TC018='4' THEN '免稅 'WHEN TC018='9' THEN '不計稅' END AS TC018NAME
                                ,PURTC.UDF02 AS 'UOF單號'
                                ,(SELECT TOP 1 [COMMENT] FROM [192.168.1.223].[UOF].[dbo].[TB_WKF_TASK_PUR_COMMENT] WHERE [TB_WKF_TASK_PUR_COMMENT].[DOC_NBR]=PURTC.UDF02 COLLATE Chinese_Taiwan_Stroke_BIN) AS '採購簽核意見'
                                ,[PACKAGE_SPEC] AS '外包裝及驗收標準'
                                ,[PRODUCT_APPEARANCE] AS '產品外觀'
                                ,[COLOR] AS '色澤'
                                ,[FLAVOR] AS '風味'
                                ,[BATCHNO] AS '產品批號'
                                ,[TB012] AS '請購單身備註'
                                ,INVMB.MB113
                                ,INVMA.[MA003] AS '產地'
                                ,PURMA.MA003 AS '公司全名'
                                ,PURMA.MA013 AS '聯絡人'

                                FROM [DY].dbo.PURTC,[DY].dbo.PURTD
                                LEFT JOIN  [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS] ON [TB_ORIENTS_CHECKLISTS].MB001=TD004
                                LEFT JOIN  [DY].dbo.INVMB ON INVMB.MB001=TD004
                                LEFT JOIN (SELECT MA001,MA002,MA003 FROM [TK].dbo.INVMA ) INVMA ON INVMA.MA001='7' AND INVMA.MA002=MB113
                                LEFT JOIN [DY].dbo.PURTB ON TB001=TD026 AND TB002=TD027 AND TB003=TD028
                                ,[DY].dbo.CMSMQ,[DY].dbo.PURMA,[DY].dbo.CMSMV,[DY].dbo.CMSMB
                                WHERE TC001=TD001 AND TC002=TD002
                                AND MQ001=TC001
                                AND TC004=PURMA.MA001
                                AND TC011=MV001
                                AND TC010=CMSMB.MB001
                                AND TC001+TC002 IN ({0})
                                {1}
 
                                ORDER BY TC001,TC002,TD003
                                ", PRINTSPURTCPURTD, STRQUERY.ToString());

            return FASTSQL;
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["採購單別"].Value.ToString();
                    textBox2.Text = row.Cells["採購單號"].Value.ToString();
                    textBox3.Text = row.Cells["供應廠"].Value.ToString();
                }
                else
                {
                    textBox1.Text = null;
                    textBox2.Text = null;
                    textBox3.Text = null;
                }
            }
        }

        public void PREPRINTS_FAX(string statusReports, string TC001, string TC002, string MA002, string COMMENT)
        {
            SETFASTREPORT_FAX(statusReports, TC001, TC002, MA002, COMMENT);
            //MessageBox.Show(PRINTSPURTCPURTD);
        }

        public void SETFASTREPORT_FAX(string statusReports, string TC001, string TC002, string MA002, string COMMENT)
        {
            string DirectoryNAME = null;
            string PDFFILES = null;
            string DATES = DateTime.Now.ToString("yyyyMMdd");

            PDF_PATH = SET_PDF_PATH();
            PDFFILES = @"" + PDF_PATH + @"\" + DATES.ToString() + @"\" + TC001 + TC002 + "-" + MA002 + ".pdf";

            DirectoryNAME = @"" + PDF_PATH + @"\" + DATES.ToString() + @"\";
            //如果日期資料夾不存在就新增 
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();

            report1.Load(@"REPORT\採購單憑証V6-核準-雅芳-大潁.frx");
            //傳入SIGNS，當簽名檔的判斷
            report1.SetParameterValue("SIGNS", statusReports);

            //if (statusReports.Equals("憑証回傳"))
            //{
            //    report1.Load(@"REPORT\採購單憑証V6-無核準-大潁.frx");
            //}
            //else if (statusReports.Equals("雅芳-簽名"))
            //{
            //    report1.Load(@"REPORT\採購單憑証V6-核準-雅芳-大潁.frx");
            //}
            //else if (statusReports.Equals("芳梅-簽名"))
            //{
            //    report1.Load(@"REPORT\採購單憑証-芳梅-核準V2-大潁.frx");
            //}

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            // ★★★ 2. 加長 FastReport 查詢逾時時間（例如 180 秒；設為 0 代表無限制）★★★
            report1.Dictionary.Connections[0].CommandTimeout = 180;

            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;

            SQL = SETFASETSQL_FAX(statusReports, TC001, TC002);
            report1.SetParameterValue("P1", COMMENT);

            Table.SelectCommand = SQL.ToString(); ;

            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            //FastReport.Export.Image.ImageExport ImageExport = new FastReport.Export.Image.ImageExport();
            // show the export options dialog and do the export
            report1.Export(export, PDFFILES);

            //傳真
            FAX(PDFFILES);
        }

        public StringBuilder SETFASETSQL_FAX(string statusReports, string TC001, string TC002)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            if (statusReports.Equals("有簽名"))
            {
                STRQUERY.AppendFormat(@"  
                                        AND TC014 IN ('Y')
                                        ");
            } 
            else
            {
                STRQUERY.AppendFormat(@"
                                        
                                        ");
            }

            FASTSQL.AppendFormat(@"      
                                SELECT *
                                ,CASE WHEN TC018='1' THEN '應稅內含' WHEN TC018='2' THEN '應稅外加' WHEN TC018='3' THEN '零稅率' WHEN TC018='4' THEN '免稅 'WHEN TC018='9' THEN '不計稅' END AS TC018NAME
                                ,PURTC.UDF02 AS 'UOF單號'
                                ,(SELECT TOP 1 [COMMENT] FROM [192.168.1.223].[UOF].[dbo].[TB_WKF_TASK_PUR_COMMENT] WHERE [TB_WKF_TASK_PUR_COMMENT].[DOC_NBR]=PURTC.UDF02 COLLATE Chinese_Taiwan_Stroke_BIN) AS '採購簽核意見'
                                ,[PACKAGE_SPEC] AS '外包裝及驗收標準'
                                ,[PRODUCT_APPEARANCE] AS '產品外觀'
                                ,[COLOR] AS '色澤'
                                ,[FLAVOR] AS '風味'
                                ,[BATCHNO] AS '產品批號'
                                ,[TB012] AS '請購單身備註'
                                ,INVMB.MB113
                                ,INVMA.[MA003] AS '產地'
                                ,PURMA.MA003 AS '公司全名'
                                ,PURMA.MA013 AS '聯絡人'

                                FROM [DY].dbo.PURTC,[DY].dbo.PURTD
                                LEFT JOIN  [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS] ON [TB_ORIENTS_CHECKLISTS].MB001=TD004
                                LEFT JOIN  [DY].dbo.INVMB ON INVMB.MB001=TD004
                                LEFT JOIN (SELECT MA001,MA002,MA003 FROM [DY].dbo.INVMA ) INVMA ON INVMA.MA001='7' AND INVMA.MA002=MB113
                                LEFT JOIN [DY].dbo.PURTB ON TB001=TD026 AND TB002=TD027 AND TB003=TD028
                                ,[DY].dbo.CMSMQ,[DY].dbo.PURMA,[DY].dbo.CMSMV,[DY].dbo.CMSMB
                                WHERE TC001=TD001 AND TC002=TD002
                                AND MQ001=TC001
                                AND TC004=PURMA.MA001
                                AND TC011=MV001
                                AND TC010=CMSMB.MB001
                                AND TC001='{0}' AND TC002='{1}'
                                {2}
 
                                ORDER BY TC001,TC002,TD003
                                ", TC001, TC002, STRQUERY.ToString());

            return FASTSQL;
        }

        public void FAX(string PDFFILES)
        {
            //一定要先安裝 Acrobat  Reader DC
            //string filePath = @"C:\採購單憑証-核準NAMEV2.pdf"; // 傳真的PDF文件路徑
            DataTable DT = FIND_FAX();
            string filePath = PDFFILES; // 傳真的PDF文件路徑
            string printerName = "LAN-Fax Generic"; // LAN-Fax 驅動名稱
            if (DT != null && DT.Rows.Count >= 1)
            {
                printerName = DT.Rows[0]["PRINTSNAMES"].ToString();
            }


            // 檢查文件是否存在
            if (!File.Exists(filePath))
            {
                Console.WriteLine("文件不存在：" + filePath);
                return;
            }

            // 使用 Acrobat Reader 或默認 PDF 閱讀器進行打印
            Process process = new Process();
            process.StartInfo.FileName = filePath; // 文件路徑
            process.StartInfo.Verb = "printto";   // 使用 "printto" 動詞直接打印到指定打印機
            process.StartInfo.Arguments = $"\"{printerName}\""; // 指定目標打印機
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = true;

            try
            {
                process.Start();
                process.WaitForExit(5000); // 等待最多 5 秒
                //Console.WriteLine("傳真發送完成！");
                //MessageBox.Show("傳真發送完成");
            }
            catch (Exception ex)
            {
                //Console.WriteLine($"傳真發送失敗: {ex.Message}");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                process.Close();
            }
        }

        public DataTable FIND_FAX()
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



                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                   SELECT [PRINTSNAMES],[KINDS]
                                    FROM [TKPUR].[dbo].[PRINTSNAMES]
                                    WHERE [KINDS]='FAX'

                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds.Tables["TEMPds1"];
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
        public void ADD_TBPURCHECKFAX(string TC001, string TC002)
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
                                   INSERT INTO [TKPUR].[dbo].[TBPURCHECKFAX] 
                                    (
                                    TC001,
                                    TC002
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    )
                                   
                                    ", TC001, TC002);

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

                    //MessageBox.Show("完成");
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
        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            PREPRINTS(comboBox1.Text.ToString(), textBox5.Text);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            //將已產生過的pdf，當傳真已確認，在「Z:\1300 資材部群組\傳真記錄區-待寄」
            ADD_TBPURCHECKFAX(textBox1.Text.Trim(), textBox2.Text.Trim());

            //產生侇真用的pdf、並呼傳真程式，在「Z:\1300 資材部群組\傳真記錄區-待寄」
            PREPRINTS_FAX(comboBox1.Text.ToString(), textBox1.Text.Trim(), textBox2.Text.Trim(), textBox3.Text.Trim(), textBox5.Text.Trim());

        }


        #endregion


    }
}
