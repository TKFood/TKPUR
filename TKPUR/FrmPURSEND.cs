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

            comboBox1load();
        }

        private void FrmPURSEND_Load(object sender, EventArgs e)
        {
            SETGRIDVIEW();
        }
        #region FUNCTION
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
            Sequel.AppendFormat(@"SELECT FORM FROM [TKPUR].[dbo].[PURREPORTFORM] WHERE [REPORT]='採購單' ORDER BY ID  ");
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
                                   SELECT ''AS '填寫EMAIL說明',TC001 AS '採購單別',TC002 AS '採購單號',TC003 AS '採購日期',TC004 AS '供應廠商',MA002 AS '供應廠',MA011 AS 'EMAIL'
                                    ,(      SELECT TD004+TD005+TD006+', '
                                            FROM   [TK].dbo.PURTD WHERE TD001=TC001 AND TD002=TC002
                                            FOR XML PATH(''), TYPE  
                                            ).value('.','nvarchar(max)')  As '明細' 
                                    FROM [TK].dbo.PURTC,[TK].dbo.PURMA
                                    WHERE 1=1
                                    AND TC004=MA001
                                    AND TC003>='{0}' AND TC003<='{1}'                                  
                                   
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

                        dataGridView1.Columns["填寫EMAIL說明"].Width = 200;

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

            if(comboBox1.Text.ToString().Equals("COA"))
            {
                report1.Load(@"REPORT\採購單憑証COA.frx");
            }
            else if (comboBox1.Text.ToString().Equals("進口報價"))
            {
                report1.Load(@"REPORT\採購單憑証進口報價.frx");
            }
            else if(comboBox1.Text.ToString().Equals("COA+進口報價"))
            {
                report1.Load(@"REPORT\採購單憑証COA進口報價.frx");
            }
            else 
            {
                report1.Load(@"REPORT\採購單憑証.frx");
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
                    dt.Columns.Add(new DataColumn("EMAILCOMMETS", typeof(string)));

                    DataRow NEWdr = dt.NewRow();
                    NEWdr["TC001"] = dr.Cells["採購單別"].Value.ToString();
                    NEWdr["TC002"] = dr.Cells["採購單號"].Value.ToString();
                    NEWdr["MA001"] = dr.Cells["供應廠商"].Value.ToString();
                    NEWdr["MA002"] = dr.Cells["供應廠"].Value.ToString();
                    NEWdr["MA011"] = dr.Cells["EMAIL"].Value.ToString();
                    NEWdr["EMAILCOMMETS"] = dr.Cells["填寫EMAIL說明"].Value.ToString();

                    dt.Rows.Add(NEWdr);
                    DSPURTCTD.Tables.Add(dt);

                    //產生附件pdf                    
                    SETFASTREPORT(DSPURTCTD);
                    //產生採購明細ds
                    DataSet DSMAILPURTCTD = FINDEMAILPURTCTD(dr.Cells["採購單別"].Value.ToString(), dr.Cells["採購單號"].Value.ToString());
                    //準備寄送email+寄送副件
                    PREPARESENDEMAIL("", dr.Cells["EMAIL"].Value.ToString(), MAILATTACHPATH, DSMAILPURTCTD, DSPURTCTD);

                   

                }
            }
        }


        public void PREPARESENDEMAIL(string FROMEMAIL, string TOEMAIL, string Attachments,DataSet DSMAILPURTCTD,DataSet DSPURTCTD)
        {
            string TC001 = null;
            string TC002 = null;
             
            try
            {
                StringBuilder SUBJEST = new StringBuilder();
                StringBuilder BODY = new StringBuilder();

                ////加上附圖
                //string path = System.Environment.CurrentDirectory+@"/Images/emaillogo.jpg";
                //LinkedResource res = new LinkedResource(path);
                //res.ContentId = Guid.NewGuid().ToString();

                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"老楊食品-採購單，請將附件用印回簽，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                string COMMENTS = DSPURTCTD.Tables[0].Rows[0]["EMAILCOMMETS"].ToString();
                BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br><br>" + "附件為老楊食品-採購單" + "<br>" + "請將附件用印回簽" + "<br>" + "謝謝" + "<br>"+ "<br>" + "請注意說明:" + COMMENTS + "</span><br>");


                if (DSMAILPURTCTD.Tables[0].Rows.Count > 0)
                {
                    BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "採購明細");

                    BODY.AppendFormat(@"<table> ");
                    BODY.AppendFormat(@"<tr >");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">品名</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購數量</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">採購單位</th>");
                    BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">到貨日</th>");
                    BODY.AppendFormat(@"</tr> ");

                    foreach (DataRow DR in DSMAILPURTCTD.Tables[0].Rows)
                    {
                        TC001 = DR["TC001"].ToString();
                        TC002 = DR["TC002"].ToString();

                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD005"].ToString() +" </td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD008"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD009"].ToString() + "</td>");
                        BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["TD012"].ToString() + "</td>");
                        BODY.AppendFormat(@"</tr> ");

                        //BODY.AppendFormat("<span></span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br> " + "品名     " + DR["TD005"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購數量 " + DR["TD008"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>" + "採購單位 " + DR["TD009"].ToString() + "</span>");
                        //BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體' > <br>");
                    }
                    BODY.AppendFormat(@"</table> ");
                }

                BODY.AppendFormat(@"

                                    ");

               

                //寄給廠商
                SENDMAIL(SUBJEST, BODY, FROMEMAIL, TOEMAIL, Attachments);
                
                //寄送副件給採購
                SENDMAILPURCC(SUBJEST, BODY, FROMEMAIL, TOEMAIL, Attachments);
            }
            catch
            {
                MessageBox.Show("有錯誤"+ TC001 + TC002);
            }
            finally
            {

            }
            
        }

        public void SENDMAIL(StringBuilder Subject, StringBuilder Body,string FROMEMAIL, string TOEMAIL, string Attachments)
        {
            try
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
                MyMail.IsBodyHtml = true; //是否使用html格式

                //加上附圖
                string path = System.Environment.CurrentDirectory + @"/Images/emaillogo.jpg";
                MyMail.AlternateViews.Add(GetEmbeddedImage(path, Body));

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
                    MessageBox.Show("有錯誤");

                    //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                    //ex.ToString();
                }
            }

            catch
            {
                MessageBox.Show("有錯誤");
            }

            finally
            {

            }
            
        }
        private AlternateView GetEmbeddedImage(String filePath,StringBuilder BODY)
        {
            LinkedResource res = new LinkedResource(filePath);
            res.ContentId = Guid.NewGuid().ToString();
            //string htmlBody = @"<img src='cid:" + res.ContentId + @"'/>";
            StringBuilder htmlBody = new StringBuilder();
            htmlBody.AppendFormat(@"
                                    <span style='font-family:新細明體,serif;color:#1F497D'><br>★☆★☆★☆★☆★☆★☆★☆★★☆★☆★☆★☆★☆★☆★☆★</span>
                                    <span lang=EN-US style='color:#1F497D'><br></span>
                                    <span style='font-size:14.0pt;font-family:標楷體'><br>資材部 徐雅芳
                                    <br>
                                    <span lang=EN-US style='font-size:14.0pt;font-family:標楷體'><br>Tel 886-5-2956520 #2000 Fax 886-5-2956519<o:p></o:p>
                                    <br>
                                    <span style='font-size:14.0pt;font-family:標楷體'><br>地址：嘉義縣大林鎮大埔美園區五路號
                                    <br>
                                    <br>
                                    <br>
                                    <span style='font-size:14.0pt;font-family:標楷體'>官網：
                                    <span lang=EN-US><a href=""http://www.tkfood.com.tw/"">
                                    <span style = 'color:blue' > http://www.tkfood.com.tw/ </span></a>
                                    <br>
                                    <span style='font-size:14.0pt;font-family:標楷體'>ＦＢ：
                                    <span lang=EN-US><a href=""https://www.facebook.com/tkfood"">
                                    <span styl ='color:blue'> https://www.facebook.com/tkfood </span></span></a>
                                    <br>
                                    <br>
                                    ");
            htmlBody.AppendFormat(@"<img src='cid:" + res.ContentId + @"'/> ");
            htmlBody.AppendFormat(@"
                                    <br>
                                    <br>
                                    <span style = 'font-size:9.0pt;font-family:標楷體'> 本電子郵件及附件所載訊息均為保密資訊，受合約保護或依法不得洩漏。其內容僅供指定收件人按限定範圍或特殊目的使用。未經授權者收到此資訊者均無權閱讀、 使用、 複製、洩漏或散佈。
                                    <br>
                                    <br>
                                    <span style = 'font-size:9.0pt;font-family:標楷體'> 老楊食品股公司將依個人資料保護法之要求妥善保管您的個人資料，並於合法取得之前提下善意使用，據此本公司僅在營運範圍內之目的與您聯繫，包含本公司主辦或協辦之行銷活動、客戶服務等，非經由本公司上開目的下之合法授權，所寄發之資訊並不代表本公司，倘若有前述情形或信件誤遞至您的信箱，請透過下列聯絡方式更正；
                                    <br>
                                    <br>
                                    <span style = 'font-size:9.0pt;font-family:標楷體'> 客服電話：
                                    <span lang = EN - US > 0800 - 522 - 109 </span>；個資服務信箱：</span>
                                    <span lang = EN - US style = 'font-family:標楷體'>
                                    <a href = ""mailto:tk100@tkfood.com.tw"">
                                    <span style = 'font-size:9.0pt;color:#0563C1'> tk100@tkfood.com.tw  </span ></a></span>
                                    <span style = 'font-size:9.0pt;font-family:標楷體'>  若您因為誤傳而收到本郵件或者非本郵件之指定收件人，煩請即刻回覆郵件告知並永久刪除此郵件及其附件和銷毀所有複印件。謝謝您的合作！</span>

                                    ");

            AlternateView alternateView = AlternateView.CreateAlternateViewFromString(BODY + htmlBody.ToString(), null, MediaTypeNames.Text.Html);
            alternateView.LinkedResources.Add(res);
            return alternateView;
        }

        public void SENDMAILPURCC(StringBuilder Subject, StringBuilder Body, string FROMEMAIL, string TOEMAIL, string Attachments)
        {
            try
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
                MyMail.Subject = "副件-" + Subject.ToString();
                //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
                MyMail.Body = Body.ToString();
                MyMail.IsBodyHtml = true; //是否使用html格式

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

                    if (DSFROMEMAIL.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow DR in DSFROMEMAIL.Tables[0].Rows)
                        {
                            MyMail.To.Add(DR["FROMEMAIL"].ToString()); //設定收件者Email，多筆mail
                                                                       //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                        }

                        MySMTP.Send(MyMail);

                        MyMail.Dispose(); //釋放資源
                    }



                }
                catch (Exception ex)
                {
                    MessageBox.Show("有錯誤");

                    //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                    //ex.ToString();
                }
            }

            catch
            {
                MessageBox.Show("有錯誤");
            }

            finally
            {

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
                                    SELECT  [FROMEMAIL] FROM [TKPUR].[dbo].[FROMEMAIL]

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
                                    AND TC001='{0}' AND TC002='{1}'
                                   ", TC001,TC002);

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

            MessageBox.Show("完成");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }

        #endregion

       
    }
}
