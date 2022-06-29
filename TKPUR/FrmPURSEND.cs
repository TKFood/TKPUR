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

namespace TKPUR
{
    public partial class FrmPURSEND : Form
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

        public FrmPURSEND()
        {
            InitializeComponent();


        }

        private void FrmPURSEND_Load(object sender, EventArgs e)
        {
            SETGRIDVIEW();
        }
        #region FUNCTION
        public void  SETGRIDVIEW()
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

        public void Search(string SDAY,string EDAY)
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
                                    
                                    AND TC002='20220629001'     
                                    ORDER BY TC001,TC002

                                    ", SDAY,EDAY);

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


        public void SETFASTREPORT(DataSet REPORTDSPURTCTD)
        {

            StringBuilder SQL=new StringBuilder();
            report1 = new Report();
            report1.Load(@"REPORT\採購單憑証.frx");

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

            if(REPORTDSPURTCTD.Tables[0].Rows.Count>0)
            {
                foreach(DataRow dr in REPORTDSPURTCTD.Tables[0].Rows)
                {
                    string MTC001 = dr["TC001"].ToString();
                    string MTC002 = dr["TC002"].ToString();
                    string MA001 = dr["MA001"].ToString();
                    string MA002 = dr["MA002"].ToString();
                    string MA011 = dr["MA011"].ToString();

                    SQL = SETFASETSQL(MTC001, MTC002);

                    Table.SelectCommand = SQL.ToString(); ;

                    //report1.Preview = previewControl1;
                    //report1.Show();

                    //// prepare a report

                    report1.PrintSettings.ShowDialog = false;
                    report1.Prepare();    // show progress dialog
                    using (var ms = new MemoryStream())
                    {
                        var pdfExport = new PDFExport
                        {
                            Name = MTC001+ MTC002+"-"+MA002+".pdf",
                            Background = true
                        };

                        report1.Export(pdfExport, ms);

                        //設定本機資料夾 
                        string DirectoryNAME = SETPATHFLODER();
                        File.WriteAllBytes(DirectoryNAME + pdfExport.Name, ms.ToArray());

                        //MEAIL附件的路徑
                        MAILATTACHPATH = null;
                        MAILATTACHPATH = DirectoryNAME + pdfExport.Name;
                    }
                }
                
            }
            


        }

        public StringBuilder SETFASETSQL(string TC001,string TC002)
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
                                AND TD001='{0}' AND TD002='{1}'
                                ", TC001, TC002);

            return FASTSQL;
        }

        //設定本機資料夾 
        public string  SETPATHFLODER()
        {
            string DirectoryNAME = null;
            string DATES = DateTime.Now.ToString("yyyyMMdd");

            DirectoryNAME = @"C:\PDFTEMP\" + DATES.ToString() + @"\";

            //不存在就新增資料夾
            if (Directory.Exists(DirectoryNAME))
            {                

            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }

            return DirectoryNAME;
        }

        public void PRESEND()
        {
            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (Convert.ToBoolean(dr.Cells[0].Value)==true)
                {
                    //MessageBox.Show(dr.Cells["採購單別"].Value.ToString()+ dr.Cells["採購單號"].Value.ToString()+ dr.Cells["EMAIL"].Value.ToString());


                    //將勾選的採購單+廠商email存成ds
                    DSPURTCTD.Clear();
                    DSPURTCTD.Reset();

                     DataTable dt = new DataTable("MyTable");
                    dt.Columns.Add(new DataColumn("TC001", typeof(string)));
                    dt.Columns.Add(new DataColumn("TC002", typeof(string)));
                    dt.Columns.Add(new DataColumn("MA001", typeof(string)));
                    dt.Columns.Add(new DataColumn("MA002", typeof(string)));
                    dt.Columns.Add(new DataColumn("MA011", typeof(string)));

                    DataRow NEWdr = dt.NewRow();
                    NEWdr["TC001"] = dr.Cells["採購單別"].Value.ToString();
                    NEWdr["TC002"] = dr.Cells["採購單號"].Value.ToString();
                    NEWdr["MA001"] = dr.Cells["供應廠商"].Value.ToString();
                    NEWdr["MA002"] = dr.Cells["供應廠"].Value.ToString();
                    NEWdr["MA011"] = dr.Cells["EMAIL"].Value.ToString();

                    dt.Rows.Add(NEWdr);
                    DSPURTCTD.Tables.Add(dt);

                    //產生附件pdf                    
                    SETFASTREPORT(DSPURTCTD);
                    //產生採購明細ds
                    DataSet DSMAILPURTCTD = FINDEMAILPURTCTD(dr.Cells["採購單別"].Value.ToString(), dr.Cells["採購單號"].Value.ToString());
                    //準備寄送email
                    PREPARESENDEMAIL("", dr.Cells["EMAIL"].Value.ToString(), MAILATTACHPATH, DSMAILPURTCTD);

                }
            }
        }


        public void PREPARESENDEMAIL(string FROMEMAIL, string TOEMAIL, string Attachments,DataSet DSMAILPURTCTD)
        {
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SUBJEST.Clear();
            BODY.Clear();
            SUBJEST.AppendFormat(@"老楊食品-採購單，請將附件用印回簽，謝謝。 "+DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
            BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為楊食品-採購單" + Environment.NewLine+"請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

            if(DSMAILPURTCTD.Tables[0].Rows.Count>0)
            {
                BODY.AppendFormat(Environment.NewLine + "採購明細");

                foreach (DataRow DR in DSMAILPURTCTD.Tables[0].Rows)
                {
                    BODY.AppendFormat(Environment.NewLine + "品名     " + DR["TD005"].ToString());
                    BODY.AppendFormat(Environment.NewLine + "採購數量 " + DR["TD008"].ToString());
                    BODY.AppendFormat(Environment.NewLine + "採購單位 " + DR["TD009"].ToString());
                    BODY.AppendFormat(Environment.NewLine );
                }
               
            }

            SENDMAIL(SUBJEST, BODY, FROMEMAIL, TOEMAIL, Attachments);
        }

        public void SENDMAIL(StringBuilder Subject, StringBuilder Body,string FROMEMAIL, string TOEMAIL, string Attachments)
        {
            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];
            DataSet DSFROMEMAIL = FINDFROMEMAIL();
            FROMEMAIL = DSFROMEMAIL.Tables[0].Rows[0]["FROMEMAIL"].ToString();

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress(FROMEMAIL);

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = Subject.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = Body.ToString();
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment(Attachments);
            MyMail.Attachments.Add(attch);

            //if (Directory.Exists(DirectoryNAME))
            //{
            //    tempFile = Directory.GetFiles(DirectoryNAME);//取得資料夾下所有檔案

            //    foreach (string item in tempFile)
            //    {
            //        info = new FileInfo(item);
            //        tFileName = info.Name.ToString().Trim();//取得檔名
            //        Attachment attch = new Attachment(DirectoryNAME+tFileName);
            //        MyMail.Attachments.Add(attch);

            //    }

            //}


            try
            {
                MyMail.To.Add(TOEMAIL); //設定收件者Email，多筆mail
                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                //ex.ToString();
            }
        }

        public DataSet FINDFROMEMAIL()
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
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT TOP 1 [FROMEMAIL] FROM [TKPUR].[dbo].[FROMEMAIL]

                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count > 0)
                {
                    return ds;
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

        public DataSet FINDEMAILPURTCTD(string TC001,string TC002)
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
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [TC001]
                                    ,[TC002]
                                    ,[TC003]
                                    ,[TC004]
                                    ,[TC005]
                                    ,[TC006]
                                    ,[TC007]
                                    ,[TC008]
                                    ,[TC009]
                                    ,[TC010]
                                    ,[TC011]
                                    ,[TC012]
                                    ,[TC013]
                                    ,[TC014]
                                    ,[TC015]
                                    ,[TC016]
                                    ,[TC017]
                                    ,[TC018]
                                    ,[TC019]
                                    ,[TC020]
                                    ,[TC021]
                                    ,[TC022]
                                    ,[TC023]
                                    ,[TC024]
                                    ,[TC025]
                                    ,[TC026]
                                    ,[TC027]
                                    ,[TC028]
                                    ,[TC029]
                                    ,[TC030]
                                    ,[TC031]
                                    ,[TC032]
                                    ,[TC033]
                                    ,[TC034]
                                    ,[TC035]
                                    ,[TC036]
                                    ,[TC037]
                                    ,[TC038]
                                    ,[TC039]
                                    ,[TC040]
                                    ,[TC041]
                                    ,[TC042]
                                    ,[TC043]
                                    ,[TC044]
                                    ,[TC045]
                                    ,[TC046]
                                    ,[TC047]
                                    ,[TC048]
                                    ,[TC049]
                                    ,[TC050]
                                    ,[TC051]
                                    ,[TC052]
                                    ,[TC053]
                                    ,[TC054]
                                    ,[TC055]
                                    ,[TC056]
                                    ,[TC057]
                                    ,[TC058]
                                    ,[TC059]
                                    ,[TC060]
                                    ,[TC061]
                                    ,[TC062]
                                    ,[TC063]
                                    ,[TC064]
                                    ,[TC065]
                                    ,[TC066]
                                    ,[TC067]
                                    ,[TC068]
                                    ,[TC069]
                                    ,[TC070]
                                    ,[TC071]
                                    ,[TC072]
                                    ,[TC073]
                                    ,[TC074]
                                    ,[TC075]
                                    ,[TC076]
                                    ,[TC077]
                                    ,[TC078]
                                    ,[TC079]
                                    ,[TC080]
                                    ,[TD001]
                                    ,[TD002]
                                    ,[TD003]
                                    ,[TD004]
                                    ,[TD005]
                                    ,[TD006]
                                    ,[TD007]
                                    ,[TD008]
                                    ,[TD009]
                                    ,[TD010]
                                    ,[TD011]
                                    ,[TD012]
                                    ,[TD013]
                                    ,[TD014]
                                    ,[TD015]
                                    ,[TD016]
                                    ,[TD017]
                                    ,[TD018]
                                    ,[TD019]
                                    ,[TD020]
                                    ,[TD021]
                                    ,[TD022]
                                    ,[TD023]
                                    ,[TD024]
                                    ,[TD025]
                                    ,[TD026]
                                    ,[TD027]
                                    ,[TD028]
                                    ,[TD029]
                                    ,[TD030]
                                    ,[TD031]
                                    ,[TD032]
                                    ,[TD033]
                                    ,[TD034]
                                    ,[TD035]
                                    ,[TD036]
                                    ,[TD037]
                                    ,[TD038]
                                    ,[TD039]
                                    ,[TD040]
                                    ,[TD041]
                                    ,[TD042]
                                    ,[TD043]
                                    ,[TD044]
                                    ,[TD045]
                                    ,[TD046]
                                    ,[TD047]
                                    ,[TD048]
                                    ,[TD049]
                                    ,[TD050]
                                    ,[TD051]
                                    ,[TD052]
                                    ,[TD053]
                                    ,[TD054]
                                    ,[TD055]
                                    ,[TD056]
                                    ,[TD057]
                                    ,[TD058]
                                    ,[TD059]
                                    ,[TD060]
                                    ,[TD061]
                                    ,[TD062]
                                    ,[TD063]
                                    ,[TD064]
                                    ,[TD065]
                                    ,[TD066]
                                    ,[TD067]
                                    ,[TD068]
                                    ,[TD069]
                                    ,[TD070]
                                    ,[TD071]
                                    ,[TD072]
                                    ,[TD073]
                                    ,[TD074]
                                    ,[TD075]
                                    ,[TD076]
                                    ,[TD077]
                                    ,[TD078]
                                    ,[TD079]
                                    ,[TD080]
                                    ,[TD081]
                                    ,[TD082]
                                    ,[TD083]
                                    ,[TD084]
                                    ,[TD085]
                                    ,[TD086]
                                    ,[TD087]
                                    ,[TD088]
                                    ,[TD089]
                                    ,[TD090]
                                    ,[TD091]
                                    ,[TD092]
                                    ,[TD093]
                                    ,[TD094]
                                    ,[TD095]
                                    ,CASE WHEN TC018='1' THEN '應稅內含' WHEN TC018='2' THEN '應稅外加' WHEN TC018='3' THEN '零稅率' WHEN TC018='4' THEN '免稅 'WHEN TC018='9' THEN '不計稅' END AS TC018NAME
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURTD,[TK].dbo.CMSMQ,[TK].dbo.PURMA,[TK].dbo.CMSMV,[TK].dbo.CMSMB
                                    WHERE TC001=TD001 AND TC002=TD002
                                    AND MQ001=TC001
                                    AND TC004=MA001
                                    AND TC011=MV001
                                    AND TC010=MB001
                                    AND TD001='A339' AND TD002='20220629001'
                                   ");

                adapter = new SqlDataAdapter(@"" + sbSql.ToString(), sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count > 0)
                {
                    return ds;
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //SETFASTREPORT();

            PRESEND();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion

       
    }
}
