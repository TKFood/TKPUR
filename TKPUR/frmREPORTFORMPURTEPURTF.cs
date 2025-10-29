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

namespace TKPUR 
{
    public partial class frmREPORTFORMPURTEPURTF : Form
    {
        int TIMEOUT_LIMITS = 240;

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

       
        public frmREPORTFORMPURTEPURTF()
        {
            InitializeComponent();

            comboBox1load();
        }


        #region FUNCTION

        private void frmREPORTFORMPURTEPURTF_Load(object sender, EventArgs e)
        {
            SETGRIDVIEW();
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
                                    (SELECT COUNT(TF005) FROM [TK].dbo.PURTF WHERE TF001=TE001 AND TF002=TE002 AND TF003=TE003) AS '明細筆數'
                                    ,TE001 AS '採購變更單別',TE002 AS '採購變更單號',TE003 AS '版次',TC003 AS '單據日期',TE004 AS '變更日期',TE005 AS '供應廠商',MA002 AS '供應廠',MA011 AS 'EMAIL'
                                    ,(      SELECT TF005+TF006+TF007+', '
                                        FROM   [TK].dbo.PURTF WHERE TF001=TE001 AND TF002=TE002 AND TF003=TE003
                                        FOR XML PATH(''), TYPE  
                                        ).value('.','nvarchar(max)')  As '明細'                                     
                                    ,(SELECT TOP 1 [COMMENT] FROM [192.168.1.223].[UOF].[dbo].[View_TB_WKF_TASK_PUR_COMMENT] WHERE [View_TB_WKF_TASK_PUR_COMMENT].[DOC_NBR]=PURTE.UDF02 COLLATE Chinese_Taiwan_Stroke_BIN) AS '採購簽核意見'

                                    FROM[TK].dbo.PURMA, [TK].dbo.PURTE
									LEFT JOIN [TK].dbo.PURTC ON TC001=TE001 AND TC002=TE002

                                    WHERE 1=1
                                    AND TE005=MA001
                                    AND TC003>='{0}' AND TC003<='{1}'                                  
                                   
                                    ORDER BY TE001,TE002,TE003

                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
        public void PREPRINTS(string statusReports,string COMMENT)
        {
            string PRINTSPURTCPURTD = null;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(dr.Cells[0].Value) == true)
                {
                    //MessageBox.Show(dr.Cells["採購單別"].Value.ToString()+ dr.Cells["採購單號"].Value.ToString());

                    PRINTSPURTCPURTD = PRINTSPURTCPURTD + "'" + dr.Cells["採購變更單別"].Value.ToString().Trim() + dr.Cells["採購變更單號"].Value.ToString().Trim() + dr.Cells["版次"].Value.ToString().Trim() + "',";
                } 
            } 

            PRINTSPURTCPURTD = PRINTSPURTCPURTD + "'A'";

            SETFASTREPORT(statusReports, PRINTSPURTCPURTD, COMMENT);
            //MessageBox.Show(PRINTSPURTCPURTD);
        }

        public void SETFASTREPORT(string statusReports, string PRINTSPURTCPURTD,string COMMENT)
        {
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();
              
            if (statusReports.Equals("憑証回傳"))
            {
                report1.Load(@"REPORT\採購單變更憑証V6-無核準.frx");
            }           
            else if (statusReports.Equals("雅芳-簽名"))
            {
                report1.Load(@"REPORT\採購單變更憑証V6-核準-雅芳.frx");
            }
            //else if (statusReports.Equals("芳梅-簽名"))
            //{
            //    report1.Load(@"REPORT\採購單變更憑証-芳梅-核準V2.frx");
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

            SQL = SETFASETSQL(statusReports, PRINTSPURTCPURTD);

            Table.SelectCommand = SQL.ToString(); ;
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
                                        AND TE017 IN ('Y')
                                        ");
            }
            else
            {
                STRQUERY.AppendFormat(@"
                                        
                                        ");
            }

            FASTSQL.AppendFormat(@"      
                                SELECT *
                                ,CASE WHEN TE018='1' THEN '應稅內含' WHEN TE018='2' THEN '應稅外加' WHEN TE018='3' THEN '零稅率' WHEN TE018='4' THEN '免稅 'WHEN TE018='9' THEN '不計稅' END AS TE018NAME
                                ,CASE WHEN TE118='1' THEN '應稅內含' WHEN TE118='2' THEN '應稅外加' WHEN TE118='3' THEN '零稅率' WHEN TE118='4' THEN '免稅 'WHEN TE118='9' THEN '不計稅' END AS TE118NAME
                                ,PURTE.UDF02 AS 'UOF單號'
                                ,CONVERT(DECIMAL(16,3),TE008) AS NEWTE008
                                ,CONVERT(DECIMAL(16,3),TE108) AS NEWTE108
                                ,CONVERT(DECIMAL(16,3),TF011) AS NEWTF011
                                ,CONVERT(DECIMAL(16,0),TF012) AS NEWTF012
                                ,CONVERT(DECIMAL(16,3),TF111) AS NEWTF111
                                ,CONVERT(DECIMAL(16,0),TF112) AS NEWTF112
                                ,(SELECT TOP 1 [COMMENT] FROM [192.168.1.223].[UOF].[dbo].[View_TB_WKF_TASK_PUR_COMMENT] WHERE [View_TB_WKF_TASK_PUR_COMMENT].[DOC_NBR]=PURTE.UDF02 COLLATE Chinese_Taiwan_Stroke_BIN) AS '採購簽核意見'
                                ,[PACKAGE_SPEC] AS '外包裝及驗收標準'
                                ,[PRODUCT_APPEARANCE] AS '產品外觀'
                                ,[COLOR] AS '色澤'
                                ,[FLAVOR] AS '風味'
                                ,[BATCHNO] AS '產品批號'
                                ,[TB012] AS '請購單身備註'

                                FROM [TK].dbo.PURTF WITH(NOLOCK)
                                LEFT JOIN  [TKRESEARCH].[dbo].[TB_ORIENTS_CHECKLISTS] ON [TB_ORIENTS_CHECKLISTS].MB001=TF005
                                LEFT JOIN [TK].dbo.PURTD ON TD001=TF001 AND TD002=TF002 AND TD003=TF004
                                LEFT JOIN [TK].dbo.PURTB ON TB001=TD026 AND TB002=TD027 AND TB003=TD028
                                ,[TK].dbo.PURTE WITH(NOLOCK)
                                LEFT JOIN [TK].dbo.CMSMQ WITH(NOLOCK) ON MQ001=TE001
                                LEFT JOIN [TK].dbo.PURMA WITH(NOLOCK) ON MA001=TE005
                                LEFT JOIN [TK].dbo.PURTC WITH(NOLOCK) ON TC001=TE001 AND TC002=TE002
                                LEFT JOIN [TK].dbo.CMSMB WITH(NOLOCK) ON TC010=MB001
                                WHERE TE001=TF001 AND TE002=TF002
                                AND TE001+TE002+TE003 IN ({0})
                                {1}

                                ORDER BY TE001,TE002,TE003,TF004
                                ", PRINTSPURTCPURTD, STRQUERY.ToString());

            return FASTSQL;
        }


        #endregion 

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        private void button1_Click(object sender, EventArgs e)
        {
            PREPRINTS(comboBox1.Text.ToString(),textBox5.Text);    
        }
         
        #endregion


    }
}
