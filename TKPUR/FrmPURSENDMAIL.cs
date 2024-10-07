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
using System.Xml;
using TKITDLL;

namespace TKPUR
{
    public partial class FrmPURSENDMAIL : Form
    {
        int TIMEOUT_LIMITS = 120;
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();


        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        public FrmPURSENDMAIL()
        {
            InitializeComponent();
        }

        #region FUNCTION
        private void FrmPURSENDMAIL_Load(object sender, EventArgs e)
        {
            comboBox1load();
            ADD_UOF_DESIGN_INFROM();
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
            Sequel.AppendFormat(@"
                                
                                SELECT 
                                [ID]
                                ,[KIND]
                                ,[PARAID]
                                ,[PARANAME]
                                FROM [TKPUR].[dbo].[TBPARA]
                                WHERE [KIND]='UOF_DESIGN_INFROM'
                                ORDER BY [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("PARAID", typeof(string));
            dt.Columns.Add("PARANAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "PARANAME";
            comboBox1.DisplayMember = "PARANAME";
            sqlConn.Close();

            //comboBox1.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void SEARCH_UOF_DESIGN_INFROM(string ISMAILS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
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

                if (!string.IsNullOrEmpty(ISMAILS))
                {
                    if (ISMAILS.Equals("全部"))
                    {
                        // 如果 ISMAILS 是 "全部"，不附加任何查询条件
                        sbSqlQuery.AppendFormat(@" ");
                    }
                    else
                    {
                        // 如果 ISMAILS 不是 "全部"，添加查询条件
                        sbSqlQuery.AppendFormat(@" AND ISMAILS IN ('{0}')", ISMAILS);
                    }
                }


                sbSql.AppendFormat(@" 
                                SELECT 
                                [SUBJECT] AS '校稿項目'
                                ,[DESIGNER] AS '設計人'
                                ,[CONTENTS]  AS '內容'
                                ,[MANUFACTOR] AS '發包廠商'
                                ,[ISMAILS]  AS '是否通知'
                                ,[MAILS_DATE] AS '通知日期'
                                FROM [TKPUR].[dbo].[UOF_DESIGN_INFROM]
                                WHERE 1=1
                                {0}

                                ", sbSqlQuery.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["ds"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["ds"];
                        dataGridView1.AutoResizeColumns();

                        //// 设定 DataGridView 的宽度的 % 给 "校稿項目" 列
                        dataGridView1.Columns["校稿項目"].Width = (dataGridView1.Width * 40) / 100;
                        dataGridView1.Columns["設計人"].Width = (dataGridView1.Width * 15) / 100;
                        dataGridView1.Columns["內容"].Width = (dataGridView1.Width * 35) / 100;
                        dataGridView1.Columns["是否通知"].Width = (dataGridView1.Width * 10) / 100;

                        // 允许 "內容" 列中的文本换行
                        dataGridView1.Columns["校稿項目"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        dataGridView1.Columns["內容"].DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                        // 自动调整行高以适应多行文本
                        dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
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

        public void ADD_UOF_DESIGN_INFROM()
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

                                    INSERT INTO [TKPUR].[dbo].[UOF_DESIGN_INFROM]
                                    (
                                    [SUBJECT]
                                    ,[BOARD_NAME]
                                    ,[CREATE_DATE]
                                    ,[CONTENTS]
                                    ,[DESIGNER] 
                                    ,[ISMAILS]
                                    )

                                    SELECT SUBJECT,BOARD_NAME,CREATE_DATE,資材
                                    ,(SELECT TOP 1 NAME
                                    FROM [192.168.1.223].[UOF].[dbo].[View_SUB_TB_EIP_FORUM_ARTICLE] 
                                    WHERE GROUP_NAME LIKE '%設計%'
                                    AND [View_SUB_TB_EIP_FORUM_ARTICLE] .SUBJECT=TEMP.SUBJECT
                                    ORDER BY CREATE_DATE) AS '設計人'
                                    ,'N' AS 'ISMAILS'

                                    FROM (
                                        SELECT 
                                            TB_EIP_FORUM_BOARD.BOARD_NAME,
                                            CONVERT(NVARCHAR, TB_EIP_FORUM_TOPIC.CREATE_DATE, 112) AS CREATE_DATE,
                                            TB_EIP_FORUM_ARTICLE.SUBJECT,
                                            ISNULL(
                                                (
                                                    SELECT TOP 1
                                                        [NAME] + ':' + CHAR(13) + CHAR(10) + CONVERT(NVARCHAR, [CREATE_DATE], 112) + CHAR(13) + CHAR(10) +
                                                        REPLACE([TKPUR].dbo.udf_StripHTML([cleaned_img_content]), '&nbsp;', '')
                                                    FROM [192.168.1.223].[UOF].[dbo].[View_SUB_TB_EIP_FORUM_ARTICLE]
                                                    WHERE [View_SUB_TB_EIP_FORUM_ARTICLE].[SUBJECT] = TB_EIP_FORUM_ARTICLE.SUBJECT
                                                    AND [View_SUB_TB_EIP_FORUM_ARTICLE].[GROUP_NAME] IN 
                                                        (SELECT [DEPNAMES] 
                                                         FROM [192.168.1.223].[UOF].[dbo].[Z_UOF_FORUM_ARTICLE_DEP] 
                                                         WHERE [DEPKINDS] IN ('資材'))
                                                    ORDER BY [View_SUB_TB_EIP_FORUM_ARTICLE].[FLOORS] DESC
                                                ), ''
                                            ) AS '資材'
        
                                        FROM 
                                            [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_AREA
                                            INNER JOIN [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_BOARD ON TB_EIP_FORUM_AREA.AREA_GUID = TB_EIP_FORUM_BOARD.AREA_GUID
                                            INNER JOIN [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_TOPIC ON TB_EIP_FORUM_BOARD.BOARD_GUID = TB_EIP_FORUM_TOPIC.BOARD_GUID
                                            INNER JOIN [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_ARTICLE ON TB_EIP_FORUM_ARTICLE.TOPIC_GUID = TB_EIP_FORUM_TOPIC.TOPIC_GUID
                                            INNER JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_USER] ON TB_EB_USER.USER_GUID = TB_EIP_FORUM_ARTICLE.ANNOUNCER
                                            INNER JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] ON TB_EB_EMPL_DEP.USER_GUID = TB_EB_USER.USER_GUID AND TB_EB_EMPL_DEP.ORDERS = '0'
                                            INNER JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_GROUP] ON TB_EB_GROUP.GROUP_ID = TB_EB_EMPL_DEP.GROUP_ID

                                        WHERE 
                                            (TB_EIP_FORUM_BOARD.BOARD_NAME LIKE '%校稿%' OR TB_EIP_FORUM_BOARD.BOARD_NAME LIKE '%設計%')
                                            AND CONVERT(NVARCHAR, TB_EIP_FORUM_TOPIC.CREATE_DATE, 112) >= '20240101'

                                        GROUP BY 
                                            TB_EIP_FORUM_BOARD.BOARD_NAME,
                                            CONVERT(NVARCHAR, TB_EIP_FORUM_TOPIC.CREATE_DATE, 112),
                                            TB_EIP_FORUM_ARTICLE.SUBJECT
                                    ) AS TEMP
                                    WHERE ISNULL(資材,'')<>''
                                    AND 資材 LIKE '%廠商%'
                                    AND SUBJECT COLLATE Chinese_Taiwan_Stroke_BIN  NOT IN (
                                    SELECT SUBJECT
                                    FROM [TKPUR].[dbo].[UOF_DESIGN_INFROM]
                                    )
                                    ORDER BY 
                                        TEMP.BOARD_NAME,
                                        TEMP.CREATE_DATE,
                                        TEMP.SUBJECT
                                   

                                    UPDATE [TKPUR].[dbo].[UOF_DESIGN_INFROM]
                                    SET [UOF_DESIGN_INFROM].[CONTENTS]=TEMP2.資材
                                    FROM 
                                    (
	                                    SELECT SUBJECT,BOARD_NAME,CREATE_DATE,資材
	                                    ,(SELECT TOP 1 NAME
	                                    FROM [192.168.1.223].[UOF].[dbo].[View_SUB_TB_EIP_FORUM_ARTICLE] 
	                                    WHERE GROUP_NAME LIKE '%設計%'
	                                    AND [View_SUB_TB_EIP_FORUM_ARTICLE] .SUBJECT=TEMP.SUBJECT
	                                    ORDER BY CREATE_DATE) AS '設計人'
	                                    ,'N' AS 'ISMAILS'

	                                    FROM (
		                                    SELECT 
			                                    TB_EIP_FORUM_BOARD.BOARD_NAME,
			                                    CONVERT(NVARCHAR, TB_EIP_FORUM_TOPIC.CREATE_DATE, 112) AS CREATE_DATE,
			                                    TB_EIP_FORUM_ARTICLE.SUBJECT,
			                                    ISNULL(
				                                    (
					                                    SELECT TOP 1
						                                    [NAME] + ':' + CHAR(13) + CHAR(10) + CONVERT(NVARCHAR, [CREATE_DATE], 112) + CHAR(13) + CHAR(10) +
						                                    REPLACE([TKPUR].dbo.udf_StripHTML([cleaned_img_content]), '&nbsp;', '')
					                                    FROM [192.168.1.223].[UOF].[dbo].[View_SUB_TB_EIP_FORUM_ARTICLE]
					                                    WHERE [View_SUB_TB_EIP_FORUM_ARTICLE].[SUBJECT] = TB_EIP_FORUM_ARTICLE.SUBJECT
					                                    AND [View_SUB_TB_EIP_FORUM_ARTICLE].[GROUP_NAME] IN 
						                                    (SELECT [DEPNAMES] 
						                                     FROM [192.168.1.223].[UOF].[dbo].[Z_UOF_FORUM_ARTICLE_DEP] 
						                                     WHERE [DEPKINDS] IN ('資材'))
					                                    ORDER BY [View_SUB_TB_EIP_FORUM_ARTICLE].[FLOORS] DESC
				                                    ), ''
			                                    ) AS '資材'
        
		                                    FROM 
			                                    [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_AREA
			                                    INNER JOIN [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_BOARD ON TB_EIP_FORUM_AREA.AREA_GUID = TB_EIP_FORUM_BOARD.AREA_GUID
			                                    INNER JOIN [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_TOPIC ON TB_EIP_FORUM_BOARD.BOARD_GUID = TB_EIP_FORUM_TOPIC.BOARD_GUID
			                                    INNER JOIN [192.168.1.223].[UOF].dbo.TB_EIP_FORUM_ARTICLE ON TB_EIP_FORUM_ARTICLE.TOPIC_GUID = TB_EIP_FORUM_TOPIC.TOPIC_GUID
			                                    INNER JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_USER] ON TB_EB_USER.USER_GUID = TB_EIP_FORUM_ARTICLE.ANNOUNCER
			                                    INNER JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_EMPL_DEP] ON TB_EB_EMPL_DEP.USER_GUID = TB_EB_USER.USER_GUID AND TB_EB_EMPL_DEP.ORDERS = '0'
			                                    INNER JOIN [192.168.1.223].[UOF].[dbo].[TB_EB_GROUP] ON TB_EB_GROUP.GROUP_ID = TB_EB_EMPL_DEP.GROUP_ID

		                                    WHERE 
			                                    (TB_EIP_FORUM_BOARD.BOARD_NAME LIKE '%校稿%' OR TB_EIP_FORUM_BOARD.BOARD_NAME LIKE '%設計%')
			                                    AND CONVERT(NVARCHAR, TB_EIP_FORUM_TOPIC.CREATE_DATE, 112) >= '20240101'

		                                    GROUP BY 
			                                    TB_EIP_FORUM_BOARD.BOARD_NAME,
			                                    CONVERT(NVARCHAR, TB_EIP_FORUM_TOPIC.CREATE_DATE, 112),
			                                    TB_EIP_FORUM_ARTICLE.SUBJECT
	                                    ) AS TEMP
	                                    WHERE ISNULL(資材,'')<>''
	                                    AND 資材 LIKE '%廠商%'
	                                    AND SUBJECT COLLATE Chinese_Taiwan_Stroke_BIN   IN (
	                                    SELECT SUBJECT
	                                    FROM [TKPUR].[dbo].[UOF_DESIGN_INFROM]
	                                    )

                                    ) AS TEMP2 
                                    WHERE TEMP2.SUBJECT=[UOF_DESIGN_INFROM].SUBJECT COLLATE Chinese_Taiwan_Stroke_BIN
                                    AND  [UOF_DESIGN_INFROM].[CONTENTS]<>TEMP2.資材 COLLATE Chinese_Taiwan_Stroke_CI_AS

                                    ");

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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox1.Text = row.Cells["校稿項目"].Value.ToString();
                }
            }
                
        }

        public void SEND_MAIL(string SUBJECTS)
        {
            DataTable DT = FIND_View_SUB_TB_EIP_FORUM_ARTICLE(SUBJECTS);

            if(DT!=null && DT.Rows.Count>=1)
            {
                SEND_MAIL_TO(DT);
            }
        }

        public DataTable FIND_View_SUB_TB_EIP_FORUM_ARTICLE(string SUBJECTS)
        {
            DataTable DT = new DataTable();
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
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

                if(!string.IsNullOrEmpty(SUBJECTS))
                {
                    sbSql.AppendFormat(@" 
                                        SELECT
                                        [SUBJECT]
                                        ,[DESIGNER]
                                        ,[BOARD_NAME]
                                        ,[CREATE_DATE]
                                        ,[CONTENTS]
                                        ,[ISMAILS]
                                        ,[MAILS_DATE]
                                        ,[NAME]
                                        ,[EMAIL]
                                        FROM [TKPUR].[dbo].[UOF_DESIGN_INFROM]
                                        LEFT JOIN [TKPUR].[dbo].[UOF_DESIGN_INFROM_EMAIL] ON [UOF_DESIGN_INFROM_EMAIL].NAME=[UOF_DESIGN_INFROM].[DESIGNER] 
                                        WHERE SUBJECT ='{0}'

                                ", SUBJECTS);
                }
                

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count >= 0)
                {
                    return ds.Tables["ds"];
                }
                else
                {
                    return null;
                }
            }
            catch
            {

            }
            finally
            {

            }
                    return DT;
        }

        public void SEND_MAIL_TO(DataTable DT)
        {
            DataTable DS_EMAIL_TO_EMAIL = new DataTable();
            DataTable DT_DATAS = new DataTable();

            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();
            //指定設計人的email
            string DESIGNER_EAMIL = "";
            //設計項目
            string SUBJETCS = "["+DT.Rows[0]["SUBJECT"].ToString()+"]";

            try
            {
                DS_EMAIL_TO_EMAIL = SERACH_MAIL_UOF_DESIGN_INFROM_EMAIL_PUR();
                DT_DATAS = DT;

                if (DT_DATAS != null && DT_DATAS.Rows.Count >= 1)
                {
                    SUBJEST.Clear();
                    BODY.Clear();


                    SUBJEST.AppendFormat(@"系統通知-請查收-採購通知校稿廠商資料-"+ SUBJETCS + "，謝謝。 " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
                    //BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為老楊食品-採購單" + Environment.NewLine + "請將附件用印回簽" + Environment.NewLine + "謝謝" + Environment.NewLine);

                    //ERP 採購相關單別、單號未核準的明細
                    //
                    BODY.AppendFormat("<span style='font-size:12.0pt;font-family:微軟正黑體'> <br>" + "Dear SIR:" + "<br>"
                        + "<br>" + "系統通知-請查收採購通知校稿廠商資料，謝謝"
                        + " <br>"
                        );





                    if (DT_DATAS.Rows.Count > 0)
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "明細");

                        BODY.AppendFormat(@"<table> ");
                        BODY.AppendFormat(@"<tr >");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">校稿項目</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">設計人</th>");
                        BODY.AppendFormat(@"<th style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">內容</th>");

                        BODY.AppendFormat(@"</tr> ");

                        foreach (DataRow DR in DT_DATAS.Rows)
                        {
                            DESIGNER_EAMIL = DR["EMAIL"].ToString();

                            BODY.AppendFormat(@"<tr >");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["SUBJECT"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["DESIGNER"].ToString() + "</td>");
                            BODY.AppendFormat(@"<td style=""border: 1px solid #999;font-size:12.0pt;font-family:微軟正黑體' "">" + DR["CONTENTS"].ToString() + "</td>");
                        
                         
                            BODY.AppendFormat(@"</tr> ");


                        }
                        BODY.AppendFormat(@"</table> ");
                    }
                    else
                    {
                        BODY.AppendFormat("<span style = 'font-size:12.0pt;font-family:微軟正黑體'><br>" + "本日無資料");
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
                            foreach (DataRow DR in DS_EMAIL_TO_EMAIL.Rows)
                            {
                                MyMail.To.Add(DR["EMAIL"].ToString()); //設定收件者Email，多筆mail
                            }

                            //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email
                            MyMail.To.Add(DESIGNER_EAMIL); //設定收件者Email
                            MySMTP.Send(MyMail);

                            MyMail.Dispose(); //釋放資源

                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show("有錯誤");

                            //ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                            //ex.ToString();
                        }
                    }
                    catch
                    {

                    }
                    finally
                    {

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

        public DataTable SERACH_MAIL_UOF_DESIGN_INFROM_EMAIL_PUR()
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

                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [NAME]
                                    ,[EMAIL]
                                    FROM [TKPUR].[dbo].[UOF_DESIGN_INFROM_EMAIL_PUR]
                                                                       
                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                // 設置查詢的超時時間，以秒為單位
                adapter.SelectCommand.CommandTimeout = TIMEOUT_LIMITS;
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
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        public void UPDATE_UOF_DESIGN_INFROM_ISMAILS(string SUBJECT)
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
                                    UPDATE [TKPUR].[dbo].[UOF_DESIGN_INFROM]
                                    SET [ISMAILS]='Y',[MAILS_DATE]=CONVERT(NVARCHAR,GETDATE(),112)
                                    WHERE [SUBJECT]='{0}'
                                    ", SUBJECT);

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
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_UOF_DESIGN_INFROM(comboBox1.Text);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADD_UOF_DESIGN_INFROM();

            SEARCH_UOF_DESIGN_INFROM(comboBox1.Text);
            MessageBox.Show("完成");

        }
        private void button3_Click(object sender, EventArgs e)
        {
            SEND_MAIL(textBox1.Text.Trim());
            //UPDATE_UOF_DESIGN_INFROM_ISMAILS(textBox1.Text.Trim());

            MessageBox.Show("完成");
        }



        #endregion


    }
}
