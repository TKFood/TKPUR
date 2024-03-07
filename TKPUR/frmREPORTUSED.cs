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
using System.Globalization;

namespace TKPUR
{
    public partial class frmREPORTUSED : Form
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
        int result;


        public frmREPORTUSED()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SETFASTREPORT(string SDAYS, string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            string SHOWSDAYS = SDAYS;
            string SHOWEDAYS = EDAYS;
            SDAYS = SDAYS + "01";
            EDAYS = EDAYS + "31";


            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@"                                    
                                SELECT 
                                TH004 AS '品號',MB002 AS '品名',MB003 AS '規格',MB004 AS '單位',CONVERT(INT,TH047) AS '進貨總金額',CONVERT(decimal(16,3),(TH047/LA011)) '平均進貨單價', LA011 AS '進貨數量'
                                ,ISNULL((SELECT SUM(TB005) FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB WHERE TA001=TB001 AND TA002=TB002 AND TA013='Y' AND TA003>='20230101' AND TA003<='20231231' AND TB003=TH004),0) AS '領用量'
                                ,ISNULL((SELECT SUM(TB005-TB004) FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB WHERE TA001=TB001 AND TA002=TB002 AND TA013='Y' AND TA003>='20230101' AND TA003<='20231231' AND TB003=TH004),0)  AS '超領量'
                                ,ISNULL((SELECT SUM(TB005) FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB WHERE TA001=TB001 AND TA002=TB002 AND TA013='Y' AND TA003>='20230101' AND TA003<='20231231' AND TB003=TH004)*0.03,0) AS 'BOM損秏3%'
                                FROM 
                                (
                                SELECT TH004,MB002,MB003,SUM(LA011) LA011,MB004,SUM(TH047) TH047
                                FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.INVLA,[TK].dbo.INVMB
                                WHERE TG001=TH001 AND TG002=TH002
                                AND LA006=TH001 AND LA007=TH002 AND LA008=TH003
                                AND TH004=MB001
                                AND TG003>='{0}' AND TG003<='{1}'
                                AND TH004 LIKE '2%'
                                GROUP BY TH004,MB002,MB003,MB004
                                ) AS TEMP
                                ORDER BY MB002,MB003,MB004
                                    ",SDAYS,EDAYS );
            SQL1 = sbSql;

            Report report1 = new Report();
            report1.Load(@"REPORT\物料使用及損秏.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SHOWSDAYS);
            report1.SetParameterValue("P2", SHOWEDAYS);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public void SETFASTREPORTBOM(string SDAYS, string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder sbSql = new StringBuilder();
            StringBuilder sbSqlQuery1 = new StringBuilder();
            string SHOWSDAYS = SDAYS;
            string SHOWEDAYS = EDAYS;
            SDAYS = SDAYS + "01";
            EDAYS = EDAYS + "31";


            sbSql.Clear();
            sbSqlQuery.Clear();


            ///BOM表用量，因BOM表修改、同時使用新舊品號，單用品號計算可能會有少算，所以要比對製令超領量
            sbSql.AppendFormat(@"                                   
                                
                                SELECT TB003 AS '品號',MB002 AS '品名',MB003 AS '規格',MB004 AS '單位',SUM(TB004) AS '總製令需量',SUM(TB005) AS '總領料量',SUM(BOMUSED) AS '依成品入庫量反推領用量',SUM(OVERUSED) AS '反推領用量-實際領用量(+是多領、-是少領)',(SUM(TB005)-SUM(TB004)) AS '總超領料量'
                                ,ISNULL((SELECT SUM(LA011) FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.INVLA WHERE TG001=TH001 AND TG002=TH002 AND LA006=TH001 AND LA007=TH002 AND LA008=TH003 AND TG013='Y' AND TH004=TB003 AND TG003>='{0}' AND TG003<='{1}' ),0) AS '總進貨量'
                                ,ISNULL((SELECT SUM(TH047) FROM [TK].dbo.PURTG,[TK].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002  AND TG013='Y' AND TH004=TB003 AND TG003>='{0}' AND TG003<='{1}' ),0) AS '總進貨金額'
                                ,ISNULL((SELECT SUM(TH047) FROM [TK].dbo.PURTG,[TK].dbo.PURTH WHERE TG001=TH001 AND TG002=TH002  AND TG013='Y' AND TH004=TB003 AND TG003>='{0}' AND TG003<='{1}' )/(SELECT SUM(LA011) FROM [TK].dbo.PURTG,[TK].dbo.PURTH,[TK].dbo.INVLA WHERE TG001=TH001 AND TG002=TH002 AND LA006=TH001 AND LA007=TH002 AND LA008=TH003 AND TG013='Y' AND TH004=TB003 AND TG003>='{0}' AND TG003<='{1}' ),0) AS '平均進貨金額'
                                FROM 
                                (
                                SELECT TA001,TA002,TA006,TA017,TB003,TB004,TB005,REALSPCTS,FAKESPCTS
                                ,ISNULL( (CASE WHEN REALSPCTS>0 THEN REALSPCTS*TA017 ELSE FAKESPCTS*TA017 END ) ,0) AS 'BOMUSED'
                                ,ISNULL( (TB005-(CASE WHEN REALSPCTS>0 THEN REALSPCTS*TA017 ELSE FAKESPCTS*TA017 END )) ,0) AS 'OVERUSED'
                                FROM 
                                (
                                SELECT TA001,TA002,TA006,TA017,TB003,TB004,TB005
                                ,(
                                SELECT AVG(1/MC004*MD006/MD007*(1+MD008))
                                FROM [TK].dbo.BOMMC,[TK].dbo.BOMMD 
                                WHERE MC001=MD001 
                                AND MC001=TA006
                                AND MD003=TB003
                                ) AS 'REALSPCTS'
                                ,(
                                SELECT AVG(1/MC004*MD006/MD007 )
                                FROM [TK].dbo.BOMMC,[TK].dbo.BOMMD 
                                WHERE MC001=MD001 
                                AND MC001=TA006
                                AND MD003 LIKE SUBSTRING(TB003,1,6)+'%'
                                ) AS 'FAKESPCTS'
                                FROM [TK].dbo.MOCTA,[TK].dbo.MOCTB
                                WHERE TA001=TB001 AND TA002=TB002
                                AND TA013='Y'
                                AND (TB003 LIKE '2%')
                                AND TA003>='{0}' AND TA003<='{1}'
                                ) AS TEMP
                                ) AS TEMP2,[TK].dbo.INVMB
                                WHERE TEMP2.TB003=MB001
                                GROUP BY TB003,MB002,MB003,MB004

                                ORDER BY MB002,TB003,MB003,MB004
                                    ", SDAYS, EDAYS);
            SQL1 = sbSql;

            Report report1 = new Report();
            report1.Load(@"REPORT\物料使用及損秏bom.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;



            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SHOWSDAYS);
            report1.SetParameterValue("P2", SHOWEDAYS);

            report1.Preview = previewControl2;
            report1.Show();
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMM"), dateTimePicker2.Value.ToString("yyyyMM"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORTBOM(dateTimePicker3.Value.ToString("yyyyMM"), dateTimePicker4.Value.ToString("yyyyMM"));
        }

        #endregion


    }
}
