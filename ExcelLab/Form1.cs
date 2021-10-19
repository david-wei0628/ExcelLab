using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace ExcelLab
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        String diolog;//路徑暫存

        /// <summary>
        /// 建檔失敗，office excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string FileStr = "D:\\Test.";
            Excel.Application Excel_App1 = new Excel.Application();
            Excel.Workbook Excel_WB = Excel_App1.Workbooks.Add();
            Excel.Worksheet Excel_WS = new Excel.Worksheet();

            Excel_WB.SaveAs(FileStr);

            //Excel_WS = null;
            //Excel_WB.Close();
            //Excel_WB = null;
            //Excel_App1.Quit();
            //Excel_App1 = null;
        }

        /// <summary>
        /// NPOI xlsx 另外選擇檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            IWorkbook workbook = null;  //新建IWorkbook物件
            string fileName = diolog;
            FileStream fileStream = new FileStream(diolog, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本
            {
                workbook = new XSSFWorkbook(fileStream);  //xlsx資料讀入workbook
            }
            else if (fileName.IndexOf(".xls") > 0) // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);  //xls資料讀入workbook
            }

            ISheet sheet = workbook.GetSheetAt(0);  //獲取第一個工作表
            IRow row;// = sheet.GetRow(0);            //新建當前工作表行資料
            //for (int i = 0; i < sheet.LastRowNum; i++)  //對工作表每一行
            //{
            //    row = sheet.GetRow(i);   //row讀入第i行資料
            //    if (row != null)
            //    {
            //        for (int j = 0; j < row.LastCellNum; j++)  //對工作表每一列
            //        {
            //            string cellValue = row.GetCell(j).ToString(); //獲取i行j列資料
            //            Console.WriteLine(cellValue);
            //        }
            //    }
            //}
            row = sheet.GetRow(1);
            //label1.Text = row.GetCell(0).ToString();

            Console.ReadLine();
            fileStream.Close();
            workbook.Close();
        }

        /// <summary>
        /// 檔案選擇 路徑與單一檔案
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialog path = new FolderBrowserDialog();
            OpenFileDialog file = new OpenFileDialog();
            //path.ShowDialog();
            file.ShowDialog();
            diolog =file.FileName;
            label1.Text = diolog;
            //OpenFileDialog fileDialog = new OpenFileDialog();
            //fileDialog.Multiselect = true;
            //fileDialog.Title = "請選擇檔案";
            //fileDialog.Filter = "所有檔案(*xls*)|*.xls*"; //設定要選擇的檔案的型別
            //if (fileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    string file = fileDialog.FileName;//返回檔案的完整路徑                
            //}
        }
    }
}
