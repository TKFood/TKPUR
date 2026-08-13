using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKPUR
{
    public partial class Frm_PURTB_PURTD_MODIFY : Form
    {
        public Frm_PURTB_PURTD_MODIFY()
        {
            InitializeComponent();
        }
        private void Frm_PURTB_PURTD_MODIFY_Load(object sender, EventArgs e)
        {
            SET_comboBox1();
        }


        #region FUNCTION
        public void SET_comboBox1()
        {
            comboBox1.Items.Clear();
            comboBox1.Items.Add("Y");
            comboBox1.Items.Add("N");
        }

        public void SEARCH_DG1(string SDATES)
        {
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery = new StringBuilder();
            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                   SELECT 
                                    TA001 AS '單別'
                                    ,TA002 AS '單號'
                                    ,TB003 AS '序號'
                                    ,TB004 AS '品號'
                                    ,TB005 AS '品名'
                                    ,TB039  AS '結案碼'
                                    FROM [TK].dbo.PURTB
                                    INNER JOIN [TK].dbo.PURTA ON TA001=TB001 AND TA002=TB002
                                    WHERE TA002 LIKE '%{0}%'
                                    ORDER BY TA001,TA002,TB003

                                    ", SDATES);
                SqlDataAdapter da = new SqlDataAdapter(@"" + sbSql, sqlConn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;


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
            SET_TEXTBOX_NULL();

            if (dataGridView1.CurrentRow != null)
            {
                string ta001 = dataGridView1.CurrentRow.Cells["單別"].Value.ToString();
                string ta002 = dataGridView1.CurrentRow.Cells["單號"].Value.ToString();
                string tb003 = dataGridView1.CurrentRow.Cells["序號"].Value.ToString();
                string tb039 = dataGridView1.CurrentRow.Cells["結案碼"].Value.ToString();

                textBox1.Text = ta001;
                textBox2.Text = ta002;
                textBox3.Text = tb003;
                comboBox1.Text = tb039;
            }

        }

        public void SET_TEXTBOX_NULL()
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";
            
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string SDATES = dateTimePicker1.Value.ToString("yyyyMMdd");            
            SEARCH_DG1(SDATES);
        }


        #endregion

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
