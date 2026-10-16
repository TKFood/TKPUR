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
                else
                {
                    sb2.AppendFormat(" AND TD016 NOT IN ('Y','y') ");
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

       
    }
}
