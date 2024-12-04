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
using System.Drawing;
using System.Drawing.Printing;
using System.Diagnostics;

namespace TKPUR
{
    public partial class FrmTEST : Form
    {
        public FrmTEST()
        {
            InitializeComponent();
        }


        #region FUNCTION


        #endregion

        #region BUTTON
        private void button2_Click(object sender, EventArgs e)
        {
            //一定要先安裝 Acrobat  Reader DC
            string filePath = @"C:\採購單憑証-核準NAMEV2.pdf"; // 傳真的PDF文件路徑
            string printerName = "LAN-Fax Generic"; // LAN-Fax 驅動名稱

            // 檢查文件是否存在
            if (!File.Exists(filePath))
            {
                Console.WriteLine("文件不存在：" + filePath);
                return;
            }

            // 使用 Acrobat Reader 或默認 PDF 閱讀器進行打印
            Process process = new Process();
            process.StartInfo.FileName = filePath; // 文件路徑
            process.StartInfo.Verb = "printto";   // 使用 "printto" 動詞直接打印到指定打印機
            process.StartInfo.Arguments = $"\"{printerName}\""; // 指定目標打印機
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.UseShellExecute = true;

            try
            {
                process.Start();
                process.WaitForExit(5000); // 等待最多 5 秒
                Console.WriteLine("傳真發送完成！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"傳真發送失敗: {ex.Message}");
            }
            finally
            {
                process.Close();
            }
        }

        #endregion
    }
}
