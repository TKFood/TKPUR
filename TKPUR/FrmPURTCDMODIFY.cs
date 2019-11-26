﻿using System;
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
    public partial class FrmPURTCDMODIFY : Form
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

        int result;
        Thread TD;

        string TD001;
        string TD002;
        string TD003;
        string TD012;
        string TD014;

        public FrmPURTCDMODIFY()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            ds.Clear();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT TD001 AS '採購單別',TD002 AS '採購單號',TD003 AS '採購序號',TD012 AS '需求日',TD014 AS '單身備註',TD004 AS '品號',TD005 AS '品名',TD006 AS '規格',TD008 AS '採購量',TD009 AS '單位'");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTD");
                sbSql.AppendFormat(@"  WHERE TD001='{0}' AND TD002='{1}'",textBox1.Text,textBox2.Text);
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
            textBox3.Text = null;
            textBox4.Text = null;
            textBox5.Text = null;
            textBox6.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];

                    textBox3.Text = row.Cells["單身備註"].Value.ToString();
                    textBox4.Text = row.Cells["採購單別"].Value.ToString();
                    textBox5.Text = row.Cells["採購單號"].Value.ToString();
                    textBox6.Text = row.Cells["採購序號"].Value.ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(row.Cells["需求日"].Value.ToString().Substring(0,4)+"/"+ row.Cells["需求日"].Value.ToString().Substring(4, 2) + "/" + row.Cells["需求日"].Value.ToString().Substring(6, 2));
                }
                else
                {
                    dataGridView1.DataSource = null;

                    textBox3.Text = null;
                    textBox4.Text = null;
                    textBox5.Text = null;
                    textBox6.Text = null;

                }
            }
        }

        public void UPDATE(string TD001,string TD002,string TD003,string TD012,string TD014)
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

                sbSql.AppendFormat(@" UPDATE [TK].dbo.PURTD ");
                sbSql.AppendFormat(@" SET TD012='{0}' ,TD014='{1}'",TD012,TD014);
                sbSql.AppendFormat(@" WHERE TD001='{0}' AND TD002='{1}' AND TD003='{2}' ",TD001,TD002,TD003);
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

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
            if(!string.IsNullOrEmpty(textBox3.Text))
            {
                UPDATE(textBox4.Text,textBox5.Text,textBox6.Text,dateTimePicker1.Value.ToString("yyyyMMdd"),textBox3.Text);

                Search();
            }
           
        }

        #endregion


    }
}
