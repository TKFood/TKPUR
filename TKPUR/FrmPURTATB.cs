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

namespace TKPUR
{
    public partial class FrmPURTATB : Form
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
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5= new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5= new DataSet();
        DataTable dt = new DataTable();
        DataTable dtADD = new DataTable();

        string tablename = null;
        string EDITID;
        int result;
        Thread TD;

        string STATUS = null;
        public Report report1 { get; private set; }
        string REPORTID;
        string DELID;

        int ROWSINDEX = 0;
        int COLUMNSINDEX = 0;

        public FrmPURTATB()
        {
            InitializeComponent();
            SETdtADD();
        }

        #region FUNCTION

        private void FrmPURTATB_Load(object sender, EventArgs e)
        {
            dataGridView4.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;      //奇數列顏色
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;


            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 120;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView4.Columns.Insert(0, cbCol);

           
            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView4.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 18;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 9);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //将 CheckBox 加入到 dataGridView
            dataGridView4.Controls.Add(cbHeader);

        }
        public void SETdtADD()
        {
            dtADD.Columns.Add("日期", typeof(String));
            dtADD.Columns.Add("請購單別", typeof(String));
            dtADD.Columns.Add("請購單號", typeof(String));
            dtADD.Columns.Add("修改次數", typeof(String));
            dtADD.Columns.Add("品號", typeof(String));
            dtADD.Columns.Add("品名", typeof(String));
            dtADD.Columns.Add("規格", typeof(String));
            dtADD.Columns.Add("單位", typeof(String));
            dtADD.Columns.Add("請購數量", typeof(String));

        }
        public void Search()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[COMMENT] AS '單頭備註', [ID] ");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002] LIKE '%{1}%' ",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[COMMENT]");
                sbSql.AppendFormat(@"  ");

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

                if (ROWSINDEX > 0 || COLUMNSINDEX > 0)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[ROWSINDEX].Cells[COLUMNSINDEX];

                    DataGridViewRow row = dataGridView1.Rows[ROWSINDEX];

                    if (ROWSINDEX >= 0)
                    {
                        DataGridViewRow DGrow = dataGridView1.Rows[ROWSINDEX];

                        textBox9.Text = DGrow.Cells["請購單別"].Value.ToString();
                        textBox10.Text = DGrow.Cells["請購單號"].Value.ToString();
                        textBox11.Text = DGrow.Cells["單頭備註"].Value.ToString();
                        textBox12.Text = DGrow.Cells["ID"].Value.ToString();
                        DELID = DGrow.Cells["ID"].Value.ToString();


                        if (!string.IsNullOrEmpty(DGrow.Cells["ID"].Value.ToString()))
                        {
                            Search2(DGrow.Cells["ID"].Value.ToString());
                        }
                    }
                    else
                    {
                        dataGridView2.DataSource = null;

                        textBox9.Text = null;
                        textBox10.Text = null;
                        textBox11.Text = null;
                        textBox12.Text = null;
                        DELID = null;

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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;

                if (dataGridView1.CurrentCell.RowIndex > 0 || dataGridView1.CurrentCell.ColumnIndex > 0)
                {                    
                    ROWSINDEX = dataGridView1.CurrentCell.RowIndex;
                    COLUMNSINDEX = dataGridView1.CurrentCell.ColumnIndex;

                    rowindex = ROWSINDEX;
                }

                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox9.Text = row.Cells["請購單別"].Value.ToString();
                    textBox10.Text = row.Cells["請購單號"].Value.ToString();
                    textBox11.Text = row.Cells["單頭備註"].Value.ToString();
                    textBox12.Text = row.Cells["ID"].Value.ToString();
                    DELID = row.Cells["ID"].Value.ToString();
                  

                    if (!string.IsNullOrEmpty(row.Cells["ID"].Value.ToString()) )
                    {
                        Search2(row.Cells["ID"].Value.ToString());
                    }
                }
                else
                {
                    dataGridView2.DataSource = null;

                    textBox9.Text = null;
                    textBox10.Text = null;
                    textBox11.Text = null;
                    textBox12.Text = null;
                    DELID = null;



                }
            }
        }

        public void Search2(string ID)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [TA001] AS '請購單別',[TA002] AS '請購單號',[TA003] AS '序號',[COMMENTD] AS '單身備註',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[NUM] AS '請購數量',TB011 AS '需求日期',[ID],[MID]");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat(@"  WHERE [MID]='{0}' ", ID);
                sbSql.AppendFormat(@"  ORDER BY [TA003]");
                sbSql.AppendFormat(@"  ");


                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds2.Tables["ds2"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds2.Tables["ds2"];
                        dataGridView2.AutoResizeColumns();


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

        public void UPDATEPURTATB(string ID,string COMMENT)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" UPDATE [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(" SET [COMMENT]='{0}' ", COMMENT);
                sbSql.AppendFormat(" WHERE[ID] ='{0}'", ID);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat("  ");

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

        public void ADDPURTATB(string TA001,string TA002,string COMMENT)
        {
            Guid Guid = Guid.NewGuid();
            string VERSIONS = GETMAXNO(TA001, TA002);
            try
            {            
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(" INSERT INTO [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat(" ([ID],[MID],[DATES],[TA001],[TA002],[TA003],[MB001],[MB002],[MB003],[MB004],[TB011],[NUM],[COMMENTD])");
                sbSql.AppendFormat(" SELECT NEWID(),'{0}',GETDATE(),PURTB.TB001,PURTB.TB002,PURTB.TB003,PURTB.TB004,PURTB.TB005,PURTB.TB006,PURTB.TB007,PURTB.TB011,CASE WHEN ISNULL([NUM],0)<>0 THEN [NUM] ELSE PURTB.TB009 END TB009,CASE WHEN ISNULL([COMMENTD],'')<>'' THEN [COMMENTD] ELSE  PURTB.TB012 END TB012", Guid.ToString());
                sbSql.AppendFormat(" FROM [TK].dbo.PURTB");
                sbSql.AppendFormat(" LEFT JOIN [TKPUR].dbo.PURTATBD ON PURTATBD.TA001=PURTB.TB001 AND PURTATBD.TA002=PURTB.TB002 AND PURTATBD.TA003=PURTB.TB003 AND MID=(SELECT TOP 1 ID FROM [TKPUR].[dbo].[PURTATB] WHERE TA001='{0}' AND TA002 ='{1}' ORDER BY VERSIONS DESC)", TA001, TA002);
                sbSql.AppendFormat(" WHERE TB001='{0}' AND TB002 ='{1}'", TA001, TA002);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" INSERT INTO [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(" ([ID],[DATES],[TA001],[TA002],[VERSIONS],[COMMENT])");
                sbSql.AppendFormat(" VALUES");
                sbSql.AppendFormat(" ('{0}',getdate(),'{1}','{2}','{3}','{4}')", Guid.ToString() , TA001, TA002, VERSIONS, COMMENT);
                sbSql.AppendFormat(" ");
            

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

        public string GETMAXNO(string TA001, string TA002)
        {
            string VERSIONS;
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds4.Clear();

                sbSql.AppendFormat(@"  SELECT ISNULL(MAX([VERSIONS]),'0') AS VERSIONS");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB] ");
                //sbSql.AppendFormat(@"  WHERE  TC001='{0}' AND TC003='{1}'", "A542","20170119");
                sbSql.AppendFormat(@"  WHERE  TA001='{0}' AND TA002='{1}'", TA001, TA002);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                    {
                        VERSIONS = SETVERSION(ds4.Tables["TEMPds4"].Rows[0]["VERSIONS"].ToString());
                        return VERSIONS;

                    }
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public string SETVERSION(string VERSIONS)
        {

            if (VERSIONS.Equals("0"))
            {
                return "1";
            }

            else
            {
                int serno = Convert.ToInt16(VERSIONS);
                serno = serno + 1;
              
                return serno.ToString();
            }
        }
       
        public void SETSTATUSFINALLY()
        {
            STATUS = null;
            

            STATUS = "EDIT";

        }

        public void SearchPURTATB()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT TB001 AS '請購單別',TB002 AS '請購單號',TB003 AS '序號',TB004 AS '品號',TB005 AS '品名',TB006 AS '規格',TB007 AS '單位',TB009 AS '請購數量' , CONVERT(NVARCHAR,TA003,112) AS '日期' ,'' AS '修改次數'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTB, [TK].dbo.PURTA");
                sbSql.AppendFormat(@"  WHERE TA001=TB001 AND TA002=TB002");
                sbSql.AppendFormat(@"  AND TB001='{0}' AND TB002='{1}' ",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(@"  ORDER BY  CONVERT(NVARCHAR,TA003,112) ,TB001,TB002");
                sbSql.AppendFormat(@"  ");

                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds3.Tables["ds3"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds3.Tables["ds3"];
                        dataGridView2.AutoResizeColumns();


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

        public void SETFASTREPORT(string REPORTID)
        {
            string SQL;
            string SQL2;
            report1 = new Report();
            report1.Load(@"REPORT\請購變更單.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            TableDataSource Table1 = report1.GetDataSource("Table1") as TableDataSource;

            SQL = SETFASETSQL(REPORTID);
            Table.SelectCommand = SQL;

            SQL2 = SETFASETSQL2(REPORTID);
            Table1.SelectCommand = SQL2;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public string SETFASETSQL(string REPORTID)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"   
                                SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[COMMENT] AS '單頭備註', [ID]
                                FROM [TKPUR].[dbo].[PURTATB]
                                WHERE ID IN ('{0}')
                                ", REPORTID); 

            return FASTSQL.ToString();
        }

        public string SETFASETSQL2(string REPORTID)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"     
                                SELECT [TA001] AS '請購單別',[TA002] AS '請購單號',[TA003] AS '序號',[COMMENTD] AS '單身備註',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[PURTATBD].[TB011] AS '需求日期',[NUM] AS '請購數量',CONVERT(NVARCHAR,[DATES],112) AS '日期',[ID],[MID]
                                ,TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號',TB009 AS '原請購數量'
                                FROM [TKPUR].[dbo].[PURTATBD]
                                LEFT JOIN [TK].dbo.PURTD ON TD026=TA001 AND TD027=TA002 AND TD028=TA003
                                LEFT JOIN [TK].dbo.PURTB ON TB001=TA001 AND TB002=TA002 AND TB003=TA003
                                WHERE [MID]='{0}'
                                ORDER BY [TA003]
                                ", REPORTID);

            return FASTSQL.ToString();
        }

        public void SETFASTREPORT2()
        {
            StringBuilder MID = new StringBuilder();
            StringBuilder ID = new StringBuilder();

            MID.Clear();
            ID.Clear();

            foreach (DataGridViewRow row in dataGridView4.Rows)
            {
                if (Convert.ToBoolean(row.Cells[0].Value))
                {
                    MID.AppendFormat(@" '{0}',", row.Cells["MID"].Value.ToString());
                    ID.AppendFormat(@" '{0}',",row.Cells["ID"].Value.ToString());
                }
            }

            MID.AppendFormat(@" '01e5cf12-ccd4-4d80-a767-0298eaed9bc2'");
            ID.AppendFormat(@" '01e5cf12-ccd4-4d80-a767-0298eaed9bc2'");

            string SQL;
            string SQL2;
            report1 = new Report();
            report1.Load(@"REPORT\請購變更單明細.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;
            TableDataSource Table1 = report1.GetDataSource("Table1") as TableDataSource;

            SQL = SETFASETSQL3(MID.ToString());
            Table.SelectCommand = SQL;

            SQL2 = SETFASETSQL4(ID.ToString());
            Table1.SelectCommand = SQL2;

            report1.Preview = previewControl2;
            report1.Show();

        }

        public string SETFASETSQL3(string REPORTID)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

            FASTSQL.AppendFormat(@"   SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[COMMENT] AS '單頭備註', [ID]");
            FASTSQL.AppendFormat(@"   FROM [TKPUR].[dbo].[PURTATB]");
            FASTSQL.AppendFormat(@"   WHERE [ID] IN ({0})", REPORTID);
            FASTSQL.AppendFormat(@"   ");

            return FASTSQL.ToString();
        }

        public string SETFASETSQL4(string REPORTID)
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();


            FASTSQL.AppendFormat(@"   
                                SELECT [TA001] AS '請購單別',[TA002] AS '請購單號',[TA003] AS '序號',[COMMENTD] AS '單身備註',[MB001] AS '品號',[MB002] AS '品名',[MB003] AS '規格',[MB004] AS '單位',[PURTATBD].[TB011] AS '需求日期',[NUM] AS '請購數量',CONVERT(NVARCHAR,[DATES],112) AS '日期',[ID],[MID]
                                 ,TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號',TB009 AS '原請購數量'
                                FROM [TKPUR].[dbo].[PURTATBD]
                                LEFT JOIN [TK].dbo.PURTD ON TD026=TA001 AND TD027=TA002 AND TD028=TA003 
                                LEFT JOIN [TK].dbo.PURTB ON TB001=TA001 AND TB002=TA002 AND TB003=TA003
                                WHERE [ID] IN ({0})
                                ORDER BY [TA001],[TA002],[TA003]                               
                                ", REPORTID);

            return FASTSQL.ToString();
        }

        public void Search3()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[DATES],112) AS '日期',[TA001] AS '請購單別',[TA002] AS '請購單號',[VERSIONS] AS '修改次數',[COMMENT] AS '單頭備註', [ID] ");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat(@"  WHERE [TA001]='{0}' AND [TA002] LIKE '%{1}%' ", textBox6.Text, textBox7.Text);
                sbSql.AppendFormat(@"  ORDER BY CONVERT(NVARCHAR,[DATES],112),[TA001],[TA002],[COMMENT]");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        dataGridView3.DataSource = ds4.Tables["ds4"];
                        dataGridView3.AutoResizeColumns();


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

        public void Search4()
        {           
            SqlDataAdapter adapter4 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();           
            DataSet ds4 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT CONVERT(NVARCHAR,[PURTATBD].[DATES],112) AS '日期',[PURTATBD].[TA001] AS '請購單別',[PURTATBD].[TA002] AS '請購單號',[PURTATBD].[TA003] AS '請購序號',[COMMENT] AS '單頭備註',[PURTATBD].[NUM] AS '數量', [PURTATBD].[COMMENTD] AS '單身備註',[PURTATBD].[MID] , [PURTATBD].[ID]  ");
                sbSql.AppendFormat(@"  FROM [TKPUR].[dbo].[PURTATB],[TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat(@"  WHERE [PURTATB].ID=[PURTATBD].MID");
                sbSql.AppendFormat(@"  AND [PURTATBD].[TA001]='{0}' AND [PURTATBD].[TA002] LIKE '%{1}%' ", textBox15.Text.Trim(), textBox16.Text.Trim());
                sbSql.AppendFormat(@"  ORDER BY [PURTATBD].[TA001],[PURTATBD].[TA002],[PURTATBD].[TA003]");
                sbSql.AppendFormat(@"  ");

                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    if (ds4.Tables["ds4"].Rows.Count >= 1)
                    {
                        dataGridView4.DataSource = ds4.Tables["ds4"];
                        dataGridView4.AutoResizeColumns();


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

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];

                    textBox3.Text = row.Cells["請購單別"].Value.ToString();
                    textBox4.Text = row.Cells["請購單號"].Value.ToString();
                    textBox5.Text = row.Cells["序號"].Value.ToString();
                    textBox8.Text = row.Cells["單身備註"].Value.ToString();
                    textBox14.Text = row.Cells["請購數量"].Value.ToString();
                    textBox13.Text = row.Cells["ID"].Value.ToString();

                    DateTime dt=Convert.ToDateTime (row.Cells["需求日期"].Value.ToString().Substring(0,4)+"/"+ row.Cells["需求日期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["需求日期"].Value.ToString().Substring(6, 2));
                    dateTimePicker1.Value = dt;


                }
                else
                {
                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox8.Text = null;
                    textBox14.Text = null;
                    textBox13.Text = null;

                }
            }
        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
           
            if (dataGridView3.CurrentRow != null)
            {
                int rowindex = dataGridView3.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView3.Rows[rowindex];

                    REPORTID = row.Cells["ID"].Value.ToString();
                }
                else
                {
                    REPORTID = null;
                  
                }
            }
        }

        public void DELPURTATB(string ID)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat("  DELETE [TKPUR].[dbo].[PURTATB]");
                sbSql.AppendFormat("  WHERE [ID]='{0}' ", ID);
                sbSql.AppendFormat("  ");
                sbSql.AppendFormat("  DELETE [TKPUR].[dbo].[PURTATBD]");
                sbSql.AppendFormat("  WHERE [MID]='{0}' ", ID);
                sbSql.AppendFormat("  ");

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

        public void UPDATEPURTATBD(string ID,string COMMENT,string NUM,string TB011)
        {
            try
            {

                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                     UPDATE [TKPUR].[dbo].[PURTATBD]
                                     SET [PURTATBD].[COMMENTD]='{0}',[PURTATBD].[NUM]={1},[PURTATBD].[TB011]={3}
                                     FROM [TK].dbo.[PURTB]
                                     WHERE [PURTATBD].[ID]='{2}'
                                     AND [PURTB].TB001=[PURTATBD].TA001  AND [PURTB].TB002=[PURTATBD].TA002  AND [PURTB].TB003=[PURTATBD].TA003 
                                     ", COMMENT, NUM, ID, TB011);

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
            Search();           
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox9.Text)&& !string.IsNullOrEmpty(textBox10.Text)&& !string.IsNullOrEmpty(textBox11.Text))
            {
                ADDPURTATB(textBox9.Text, textBox10.Text, textBox11.Text);
                Search();
            }
           
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox13.Text) && !string.IsNullOrEmpty(textBox8.Text))
            {
                UPDATEPURTATBD(textBox13.Text, textBox8.Text,textBox14.Text,dateTimePicker1.Value.ToString("yyyyMMdd"));
                Search();
                MessageBox.Show("完成");
            }


        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search3();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(REPORTID);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox12.Text) && !string.IsNullOrEmpty(textBox11.Text) )
            {
                UPDATEPURTATB(textBox12.Text, textBox11.Text);
                Search();
            }
               
        }
        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if(!string.IsNullOrEmpty(textBox12.Text))
                {
                    DELPURTATB(textBox12.Text);

                    Search();
                }               

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Search4();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }


        #endregion

      
    }
}
