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
                                ,[ISMAILS]  AS '是否通知'
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH_UOF_DESIGN_INFROM(comboBox1.Text);
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADD_UOF_DESIGN_INFROM();

            MessageBox.Show("完成");

        }


        #endregion

      
    }
}
