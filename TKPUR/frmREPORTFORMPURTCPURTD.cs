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
    public partial class frmREPORTFORMPURTCPURTD : Form
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

        public frmREPORTFORMPURTCPURTD()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION
        private void frmREPORTFORMPURTCPURTD_Load(object sender, EventArgs e)
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
                                   SELECT TC001 AS '採購單別',TC002 AS '採購單號',TC003 AS '採購日期',TC004 AS '供應廠商',MA002 AS '供應廠',MA011 AS 'EMAIL'
                                    ,(      SELECT TD004+TD005+TD006+', '
                                            FROM   [TK].dbo.PURTD WHERE TD001=TC001 AND TD002=TC002
                                            FOR XML PATH(''), TYPE  
                                            ).value('.','nvarchar(max)')  As '明細' 
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURMA
                                    WHERE 1=1
                                    AND TC004=MA001
                                    AND TC003>='{0}' AND TC003<='{1}'                                  
                                   
                                    ORDER BY TC001,TC002

                                    ", SDAY, EDAY);

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
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
        public void PREPRINTS()
        {
            string PRINTSPURTCPURTD = null;
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(dr.Cells[0].Value) == true)
                {
                    //MessageBox.Show(dr.Cells["採購單別"].Value.ToString()+ dr.Cells["採購單號"].Value.ToString());

                    PRINTSPURTCPURTD = PRINTSPURTCPURTD+"'"+ dr.Cells["採購單別"].Value.ToString().Trim() + dr.Cells["採購單號"].Value.ToString().Trim()+ "',";
                }
            }

            PRINTSPURTCPURTD = PRINTSPURTCPURTD + "'A'";

            SETFASTREPORT(PRINTSPURTCPURTD);
            //MessageBox.Show(PRINTSPURTCPURTD);
        }

        public void SETFASTREPORT(string PRINTSPURTCPURTD)
        {
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();

            if (comboBox1.Text.ToString().Equals("憑証回傳202209"))
            {
                report1.Load(@"REPORT\採購單憑証V2.frx");
            }

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

            SQL = SETFASETSQL(PRINTSPURTCPURTD);

            Table.SelectCommand = SQL.ToString(); ;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public StringBuilder SETFASETSQL(string PRINTSPURTCPURTD)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();


            FASTSQL.AppendFormat(@"    
                                SELECT *
                                ,CASE WHEN TC018='1' THEN '應稅內含' WHEN TC018='2' THEN '應稅外加' WHEN TC018='3' THEN '零稅率' WHEN TC018='4' THEN '免稅 'WHEN TC018='9' THEN '不計稅' END AS TC018NAME
                                FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.CMSMQ,[TK].dbo.PURMA,[TK].dbo.CMSMV,[TK].dbo.CMSMB
                                WHERE TC001=TD001 AND TC002=TD002
                                AND MQ001=TC001
                                AND TC004=MA001
                                AND TC011=MV001
                                AND TC010=MB001
                                AND TC001+TC002 IN ({0})
                                ", PRINTSPURTCPURTD); 

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
            PREPRINTS();
        }

        #endregion


    }
}
